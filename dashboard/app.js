/* GMAO leon: admin dashboard UI (routing, pages, detail sheet, tables, charts).
 * All calculations come from core.js (GMCore); this file only shapes and presents them. */
(function () {
  'use strict';
  const C = window.GMCore;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = C.norm;

  const S = {
    data: null, report: null, source: '', isDemo: false,
    period: '12m', from: null, to: null, zone: '', model: '',
    charts: [], asTable: new Set(), ui: {}, ctx: null, route: { page: 'apercu', param: null },
  };

  // ---------- formatting ----------
  const nf = (d) => new Intl.NumberFormat('fr-FR', { minimumFractionDigits: d, maximumFractionDigits: d });
  const fmtN = (v, d = 0) => (v == null || !isFinite(v) ? '—' : nf(d).format(v));
  const fmtPct = (v, d = 1) => (v == null || !isFinite(v) ? '—' : `${nf(d).format(v * 100)} %`);
  const fmtD = (s) => (s == null ? '—' : C.fmtDate(s));
  const fmtMoney = (v) => (v == null ? '—' : `${fmtN(v)} ${S.data.params.devise}`);
  const toDate = (s) => new Date((s - 25569) * 86400000);
  const isoFromSerial = (s) => toDate(s).toISOString().slice(0, 10);
  const longDate = (s) => toDate(s).toLocaleDateString('fr-FR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
  const monthKey = (s) => { const d = toDate(s); return d.getUTCFullYear() * 12 + d.getUTCMonth(); };
  const css = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  const icon = (name, cls) => `<i data-lucide="${name}" aria-hidden="true"${cls ? ` class="${cls}"` : ''}></i>`;
  const paintIcons = () => window.lucide && window.lucide.createIcons();
  const plural = (n, one, many) => `${fmtN(n)} ${n > 1 ? many : one}`;

  const BADGE_ICON = { good: 'circle-check', warn: 'circle-alert', bad: 'circle-x', info: 'info', neutral: 'circle-help' };
  const badge = (kind, label) => `<span class="badge ${kind}">${kind === 'neutral' ? '' : icon(BADGE_ICON[kind] || 'info')}${esc(label)}</span>`;
  const STATUS_LABEL = { good: 'Objectif atteint', warn: 'À surveiller', bad: 'Hors objectif', na: 'Non calculable' };
  const kpiBadge = (k) => badge(k.status === 'na' ? 'neutral' : k.status, STATUS_LABEL[k.status]);
  const CRIT_LEVEL = { Faible: 1, Moyenne: 2, 'Élevée': 3, Critique: 4 };
  const critHtml = (c) => {
    if (!c) return '<span class="muted">—</span>';
    const lv = CRIT_LEVEL[c] || 0;
    return `<span class="crit c${lv}"><span class="crit-bars" aria-hidden="true">${[1, 2, 3, 4].map((i) => `<i class="${i <= lv ? 'on' : ''}"></i>`).join('')}</span>${esc(c)}</span>`;
  };

  // ---------- data loading ----------
  const refDate = () => (S.data.params.refDate != null ? S.data.params.refDate : C.todaySerial());

  function setData(data, report, source, isDemo) {
    S.data = data; S.report = report; S.source = source; S.isDemo = isDemo;
    S.zone = ''; S.model = ''; S.ui = {};
    const o = C.filterOptions(data);
    $('fZone').innerHTML = `<option value="">Tous les services</option>${o.zones.map((z) => `<option>${esc(z)}</option>`).join('')}`;
    $('fModel').innerHTML = `<option value="">Tous les modèles</option>${o.models.map((m) => `<option>${esc(m)}</option>`).join('')}`;
    const range = C.dataRange(data);
    const span = range.min != null ? (refDate() - range.min) / 30.4 : 0;
    setPeriod(span > 13 ? '12m' : 'all', false);
    renderSource();
    update();
  }

  function loadDemo() {
    setData(C.demoData(), null, 'Parc de démonstration', true);
  }

  function loadBuffer(buf, name) {
    try {
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      const { data, report } = C.parseWorkbook(XLSX, wb);
      if (!data.pannes.length && !data.interventions.length && !data.equipements.length) {
        notice('warn', `« ${name} » ne contient aucune feuille reconnue (pannes, interventions, équipements…). Téléchargez le modèle Excel depuis la page Données et import pour voir les en-têtes attendus.`);
        return;
      }
      setData(data, report, name, false);
      toast(`Fichier « ${name} » importé : ${data.equipements.length} équipements, ${data.pannes.length} pannes.`);
    } catch (err) {
      notice('warn', `Impossible de lire « ${name} » : ${err.message}`);
    }
  }
  function loadFile(file) {
    const r = new FileReader();
    r.onload = (e) => loadBuffer(e.target.result, file.name);
    r.readAsArrayBuffer(file);
  }

  // ---------- period & filters ----------
  function setPeriod(p, rerender = true) {
    S.period = p;
    $('fPeriod').value = p;
    const ref = refDate();
    const range = C.dataRange(S.data);
    let from, to = ref;
    if (p === '30d') from = ref - 29;
    else if (p === '3m') from = C.addMonths(ref, -3) + 1;
    else if (p === '12m') from = C.addMonths(ref, -12) + 1;
    else if (p === 'ytd') from = C.toSerial(`${toDate(ref).getUTCFullYear()}-01-01`);
    else if (p === 'all') {
      from = range.min != null ? range.min : ref - 365;
      if (range.max != null && range.max > to) to = range.max;
    } else {
      from = S.from != null ? S.from : ref - 365;
      to = S.to != null ? S.to : ref;
    }
    S.from = from; S.to = to;
    $('fFrom').value = isoFromSerial(from);
    $('fTo').value = isoFromSerial(to);
    $('customDates').hidden = p !== 'custom';
    if (rerender) update();
  }

  // ---------- derived context ----------
  function scopeOK(code, zone) {
    const e = S.ctx ? S.ctx.eqBy[code] : null;
    if (S.zone && (zone || (e && e.zone)) !== S.zone) return false;
    if (S.model && (!e || (e.model || e.name) !== S.model)) return false;
    return true;
  }
  const isClosedOT = (o) => ['cloture', 'annule'].includes(norm(o.status));

  function assetStatus(e, ref) {
    const st = norm(e.status);
    if (st === 'reforme') return { key: 'neutral', label: 'Réformé', rank: 0 };
    const openCorr = S.data.ot.find((o) => o.code === e.code && C.typeGroup(o.type) === 'Correctif' && !isClosedOT(o));
    if (['a l arret', 'en reparation', 'hors service'].includes(st) || openCorr)
      return { key: 'bad', label: openCorr && /attente/i.test(openCorr.status) ? 'En attente de pièces' : 'En réparation', rank: 3, ot: openCorr };
    const late = S.data.preventif.some((t) => t.code === e.code && t.done == null && t.due != null && t.due < ref);
    if ((e.service && e.service.annualState === 'late') || late) return { key: 'warn', label: 'Maintenance en retard', rank: 2 };
    if (e.service && e.service.annualState === 'soon') return { key: 'info', label: 'Maintenance à prévoir', rank: 1 };
    return { key: 'good', label: 'Opérationnel', rank: 0 };
  }

  function buildCtx() {
    const f = { from: S.from, to: S.to, zone: S.zone, model: S.model };
    const res = C.compute(S.data, f);
    const eqBy = Object.fromEntries(S.data.equipements.map((e) => [e.code, e]));
    S.ctx = { res, eqBy };
    const ref = res.ref;
    const days = res.days;
    const range = C.dataRange(S.data);
    let prev = null;
    if (range.min != null && range.min < S.from && days > 0) prev = C.compute(S.data, { from: S.from - days, to: S.from - 1, ref: S.from - 1, zone: S.zone, model: S.model });
    const assets = res.fleet.map((e) => Object.assign(e, { st: assetStatus(e, ref) }));
    const ctx = Object.assign(S.ctx, { res, prev, assets, ref, eqBy, kpi: Object.fromEntries(res.kpis.map((k) => [k.key, k])) });
    ctx.actions = buildActions(ctx);
    ctx.months = monthlyExtras(ctx);
    return ctx;
  }

  function monthlyExtras(ctx) {
    const { res } = ctx;
    const inP = (s) => s != null && s >= res.from && s <= res.to;
    const interv = S.data.interventions.filter((i) => inP(i.date) && scopeOK(i.code, i.zone) && norm(i.status) !== 'annulee');
    const tasks = S.data.preventif.filter((t) => t.due != null && t.due >= res.from && t.due <= res.toEff && scopeOK(t.code, t.zone));
    return res.monthly.map((m) => {
      const mi = interv.filter((i) => monthKey(i.date) === m.key);
      const mt = tasks.filter((t) => monthKey(t.due) === m.key);
      const done = mt.filter((t) => t.done != null).length;
      return Object.assign({}, m, {
        cost: mi.reduce((a, i) => a + (i.costTotal || 0), 0),
        mo: mi.reduce((a, i) => a + (i.costMO || 0), 0),
        parts: mi.reduce((a, i) => a + (i.costParts || 0), 0),
        other: mi.reduce((a, i) => a + (i.costOther || 0), 0),
        due: mt.length, done, prevRate: mt.length ? done / mt.length : null,
      });
    });
  }

  function buildActions(ctx) {
    const out = [];
    const ref = ctx.ref;
    const label = (code) => { const e = ctx.eqBy[code]; return e ? `${code} · ${e.model || e.name}` : code; };
    const where = (code) => { const e = ctx.eqBy[code]; return e ? [e.zone, e.line].filter(Boolean).join(', ') : ''; };
    for (const a of ctx.assets) {
      if (a.st.key === 'bad' && a.st.ot) {
        const o = a.st.ot;
        out.push({ sev: 'bad', icon: 'wrench', title: `${label(a.code)} : ${a.st.label.toLowerCase()}`, sub: `${o.desc || 'Réparation'} · ${where(a.code)} · OT ${o.id} ouvert le ${fmtD(o.created)}`, href: `#/equipements/${encodeURIComponent(a.code)}`, cta: 'Voir l’appareil' });
      }
      if (a.service && a.service.annualState === 'late') {
        out.push({ sev: 'bad', icon: 'shield-alert', title: `${label(a.code)} : maintenance annuelle + CST en retard`, sub: `Échéance dépassée depuis ${plural(ref - a.service.nextAnnual, 'jour', 'jours')} (${fmtD(a.service.nextAnnual)}) · à planifier avec un technicien agréé`, href: `#/equipements/${encodeURIComponent(a.code)}`, cta: 'Voir l’appareil' });
      }
    }
    const lateByEq = new Map();
    for (const t of ctx.res.prevLate) {
      if (!lateByEq.has(t.code)) lateByEq.set(t.code, []);
      lateByEq.get(t.code).push(t);
    }
    for (const [code, ts] of lateByEq) {
      const oldest = ts.reduce((m, t) => Math.min(m, t.due), Infinity);
      out.push({ sev: 'warn', icon: 'calendar-x', title: `${label(code)} : ${plural(ts.length, 'tâche préventive en retard', 'tâches préventives en retard')}`, sub: `${ts.slice(0, 2).map((t) => t.op).join(' · ')}${ts.length > 2 ? '…' : ''} · plus ancienne : ${fmtD(oldest)}`, href: '#/preventif', cta: 'Planifier' });
    }
    for (const o of ctx.res.otOpen) {
      if (C.typeGroup(o.type) === 'Correctif') continue;
      if (o.planned != null && o.planned < ref)
        out.push({ sev: 'warn', icon: 'clipboard-list', title: `Ordre de travail ${o.id} en retard`, sub: `${label(o.code)} · ${o.desc} · prévu le ${fmtD(o.planned)}`, href: '#/ot', cta: 'Voir l’OT' });
    }
    for (const a of ctx.assets) {
      if (a.service && a.service.annualState === 'soon')
        out.push({ sev: 'info', icon: 'calendar-clock', title: `${label(a.code)} : maintenance annuelle à prévoir`, sub: `Dans ${plural(a.service.nextAnnual - ref, 'jour', 'jours')}, le ${fmtD(a.service.nextAnnual)} · réserver un technicien agréé`, href: `#/equipements/${encodeURIComponent(a.code)}`, cta: 'Voir l’appareil' });
    }
    const order = { bad: 0, warn: 1, info: 2 };
    return out.sort((a, b) => order[a.sev] - order[b.sev]);
  }

  // ---------- notices / toast / source ----------
  function notice(kind, text) {
    const el = document.createElement('div');
    el.className = `notice ${kind}`;
    el.innerHTML = `${icon(kind === 'warn' ? 'triangle-alert' : 'info')}<div>${esc(text)}</div><button class="icon-btn x" type="button" aria-label="Fermer le message">${icon('x')}</button>`;
    el.querySelector('.x').onclick = () => el.remove();
    $('notices').appendChild(el);
    paintIcons();
  }
  let toastTimer;
  function toast(text) {
    const t = $('toast');
    t.innerHTML = `${icon('circle-check')}<span>${esc(text)}</span>`;
    t.hidden = false;
    paintIcons();
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), 4000);
  }
  function renderSource() {
    const d = S.data;
    $('sourceCard').innerHTML = `${S.isDemo ? '<span class="demo-tag">Données fictives</span>' : ''}<strong>${esc(S.source)}</strong>${d.equipements.length} équipements · ${d.pannes.length} pannes`;
  }

  // ---------- tables (sortable + searchable; rows are moved, never rebuilt, on sort) ----------
  const tables = {};
  function dataTable({ id, caption, cols, rows, empty = 'Aucun élément.', href, max = 0 }) {
    tables[id] = { cols, rows, href };
    if (!rows.length) return `<div class="empty table-empty">${icon('inbox')}${esc(empty)}</div>`;
    const shown = max ? rows.slice(0, max) : rows;
    return `<div class="table-wrap" role="region" tabindex="0" aria-labelledby="${id}-cap">
      <table id="${id}">
        <caption id="${id}-cap" class="sr-only">${esc(caption)}</caption>
        <thead><tr>${cols.map((c, ci) => `<th scope="col" class="${c.n ? 'n' : ''}">${c.sort === false ? `<span>${esc(c.label)}</span>` : `<button type="button" data-col="${ci}">${esc(c.label)}${icon('chevron-down', 'sort-ico')}</button>`}</th>`).join('')}</tr></thead>
        <tbody>${shown.map((r, ri) => `<tr data-i="${ri}"${href ? ' class="clickable"' : ''}>${cols.map((c) => `<td class="${c.n ? 'n' : c.nw ? 'nw' : c.wide ? 'wide' : ''}">${c.html(r)}</td>`).join('')}</tr>`).join('')}</tbody>
      </table></div>
      <div class="table-foot"><span id="${id}-count">${plural(shown.length, 'ligne', 'lignes')}${max && rows.length > max ? ` (sur ${fmtN(rows.length)})` : ''}</span></div>`;
  }

  function wireTable(id, searchInput) {
    const table = document.getElementById(id);
    const def = tables[id];
    if (!table || !def) return;
    const tbody = table.tBodies[0];
    table.tHead.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-col]');
      if (!btn) return;
      const th = btn.parentElement;
      const col = def.cols[+btn.dataset.col];
      const dir = th.getAttribute('aria-sort') === 'ascending' ? 'descending' : 'ascending';
      for (const h of table.tHead.rows[0].cells) h.removeAttribute('aria-sort');
      th.setAttribute('aria-sort', dir);
      const val = col.sortVal || col.val || ((r) => norm(col.html(r).replace(/<[^>]+>/g, '')));
      const rowsEls = [...tbody.rows];
      rowsEls.sort((a, b) => {
        const va = val(def.rows[+a.dataset.i]), vb = val(def.rows[+b.dataset.i]);
        const na = va == null || va === '', nb = vb == null || vb === '';
        if (na || nb) return na === nb ? 0 : na ? 1 : -1;
        const c = typeof va === 'number' && typeof vb === 'number' ? va - vb : String(va).localeCompare(String(vb), 'fr', { numeric: true });
        return dir === 'ascending' ? c : -c;
      });
      tbody.append(...rowsEls);
      $('announcer').textContent = `Tableau trié par ${col.label}, ordre ${dir === 'ascending' ? 'croissant' : 'décroissant'}.`;
    });
    if (def.href) {
      tbody.addEventListener('click', (e) => {
        if (e.target.closest('a, button')) return;
        const tr = e.target.closest('tr');
        if (tr) location.hash = def.href(def.rows[+tr.dataset.i]);
      });
    }
    if (searchInput) {
      searchInput.addEventListener('input', () => {
        const q = norm(searchInput.value);
        let n = 0;
        for (const tr of tbody.rows) {
          const hit = !q || norm(tr.textContent).includes(q);
          tr.hidden = !hit;
          if (hit) n++;
        }
        $(`${id}-count`).textContent = q ? `${plural(n, 'résultat', 'résultats')} sur ${fmtN(tbody.rows.length)}` : plural(tbody.rows.length, 'ligne', 'lignes');
      });
    }
  }

  const searchBox = (id, ph) => `<label class="search">${icon('search')}<span class="sr-only">${esc(ph)}</span><input type="search" id="${id}" placeholder="${esc(ph)}" autocomplete="off"></label>`;
  const eqCell = (code) => {
    const e = S.ctx.eqBy[code];
    return `<a href="#/equipements/${encodeURIComponent(code)}"><strong>${esc(code)}</strong></a>${e ? `<span class="sub">${esc(e.model || e.name)}</span>` : ''}`;
  };

  // ---------- charts ----------
  const trunc = (s, n) => (s.length > n ? `${s.slice(0, n - 1)}…` : s);

  function chartCard(spec) {
    const asTable = S.asTable.has(spec.id);
    const hasData = spec.labels.length && spec.series.some((s) => s.data.some((v) => v != null && v !== 0));
    const legend = spec.series.length > 1 || spec.target
      ? `<div class="legend">${spec.series.map((s) => `<span><i style="background:${css(s.color)}"></i>${esc(s.label)}</span>`).join('')}${spec.target ? `<span><i class="line" style="background:${css('--ref-line')}"></i>${esc(spec.target.label)}</span>` : ''}</div>`
      : '';
    return `<section class="card chart-card${asTable ? ' as-table' : ''}" data-chart="${spec.id}" aria-labelledby="${spec.id}-t">
      <div class="card-h"><div><h2 id="${spec.id}-t">${esc(spec.title)}</h2>${spec.sub ? `<p>${esc(spec.sub)}</p>` : ''}</div>
        <div class="tools"><div class="view-toggle" role="group" aria-label="Affichage">
          <button type="button" data-view="chart" aria-pressed="${!asTable}">${icon('chart-column')}Graphique</button>
          <button type="button" data-view="table" aria-pressed="${asTable}">${icon('table')}Tableau</button></div></div></div>
      <div class="card-b">
        <div class="chart-view">${legend}<div class="chart-box${spec.tall ? ' tall' : ''}${spec.short ? ' short' : ''}">${hasData ? `<canvas role="img" aria-label="${esc(spec.title)} (valeurs disponibles via le bouton Tableau)"></canvas>` : `<div class="empty">${icon('chart-no-axes-column')}Aucune donnée pour cette période.</div>`}</div></div>
        <div class="tbl-view">${chartTableHtml(spec)}</div>
        ${spec.insight ? `<div class="insight">${icon('lightbulb')}<span>${spec.insight}</span></div>` : ''}
      </div></section>`;
  }

  function chartTableHtml(spec) {
    if (!spec.labels.length) return '<div class="empty">Aucune donnée.</div>';
    const head = [spec.catLabel || 'Libellé'].concat(spec.series.map((s) => s.label)).concat(spec.extraCols ? spec.extraCols.head : []);
    return `<div class="table-wrap" role="region" tabindex="0" aria-label="${esc(spec.title)}"><table><caption class="sr-only">${esc(spec.title)}</caption><thead><tr>${head.map((h, i) => `<th scope="col" class="${i ? 'n' : ''}"><span>${esc(h)}</span></th>`).join('')}</tr></thead><tbody>${spec.labels
      .map((l, i) => `<tr><td>${esc(l)}</td>${spec.series.map((s) => `<td class="n">${esc(spec.fmt(s.data[i]))}</td>`).join('')}${spec.extraCols ? spec.extraCols.row(i).map((c) => `<td class="n">${esc(c)}</td>`).join('') : ''}</tr>`)
      .join('')}</tbody></table></div>`;
  }

  function mountChart(spec) {
    const card = document.querySelector(`[data-chart="${spec.id}"]`);
    if (!card) return;
    card.querySelectorAll('.view-toggle button').forEach((b) =>
      b.addEventListener('click', () => {
        const table = b.dataset.view === 'table';
        table ? S.asTable.add(spec.id) : S.asTable.delete(spec.id);
        card.classList.toggle('as-table', table);
        card.querySelectorAll('.view-toggle button').forEach((x) => x.setAttribute('aria-pressed', String(x === b)));
      })
    );
    const canvas = card.querySelector('canvas');
    if (!canvas) return;
    const grid = css('--grid'), muted = css('--muted'), text = css('--text'), text2 = css('--text-2'), surface = css('--surface'), border = css('--border');
    const horizontal = spec.kind === 'hbar';
    const isLine = spec.kind === 'line';
    const datasets = spec.series.map((s, si) => {
      const color = css(s.color);
      if (isLine) return { label: s.label, data: s.data, borderColor: color, backgroundColor: color, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, pointBackgroundColor: color, pointBorderColor: surface, pointBorderWidth: 2, tension: 0, spanGaps: true };
      const last = si === spec.series.length - 1;
      return {
        label: s.label, data: s.data, backgroundColor: color, hoverBackgroundColor: color, borderColor: surface,
        borderWidth: spec.stacked ? (horizontal ? { right: 2 } : { top: 2 }) : 0,
        borderRadius: spec.stacked ? (last ? (horizontal ? { topRight: 4, bottomRight: 4 } : { topLeft: 4, topRight: 4 }) : 0) : 4,
        borderSkipped: spec.stacked ? false : 'start', maxBarThickness: horizontal ? 16 : 26, categoryPercentage: spec.grouped ? 0.7 : 0.8, barPercentage: spec.grouped ? 0.85 : 0.9,
      };
    });
    if (spec.target) datasets.push({ label: spec.target.label, data: spec.labels.map(() => spec.target.value), borderColor: css('--ref-line'), borderWidth: 1, borderDash: [4, 3], pointRadius: 0, pointHoverRadius: 0, fill: false });
    const valueAxis = {
      beginAtZero: !spec.pct, stacked: !!spec.stacked, grid: { color: grid, drawTicks: false }, border: { display: false },
      ticks: { color: muted, font: { size: 11.5 }, padding: 8, maxTicksLimit: 6, callback: (v) => (spec.pct ? fmtPct(v, 1) : fmtN(v)) },
    };
    if (spec.pct) {
      const vals = spec.series[0].data.filter((v) => v != null).concat(spec.target ? [spec.target.value] : []);
      valueAxis.min = Math.max(0, Math.floor((Math.min(...vals) - 0.003) * 200) / 200);
      valueAxis.max = 1;
    }
    const catAxis = {
      stacked: !!spec.stacked, grid: { display: false }, border: { color: grid },
      ticks: { color: text2, font: { size: 11.5 }, autoSkip: !horizontal, maxRotation: 0, callback: function (v) { return trunc(this.getLabelForValue(v), horizontal ? 30 : 12); } },
    };
    Chart.defaults.font.family = css('--font') || 'Inter, system-ui, sans-serif';
    const chart = new Chart(canvas, {
      type: isLine ? 'line' : 'bar',
      data: { labels: spec.labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        indexAxis: horizontal ? 'y' : 'x',
        interaction: { mode: 'index', intersect: false, axis: horizontal ? 'y' : 'x' },
        scales: horizontal ? { x: valueAxis, y: catAxis } : { x: catAxis, y: valueAxis },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: surface, titleColor: text, bodyColor: text2, footerColor: text2, borderColor: border, borderWidth: 1,
            padding: 12, boxPadding: 4, usePointStyle: true, titleFont: { weight: '600' }, footerFont: { weight: '500' },
            filter: (item) => !(spec.target && item.datasetIndex === spec.series.length),
            callbacks: {
              label: (item) => ` ${item.dataset.label} : ${spec.fmt(item.raw)}`,
              afterBody: (items) => (spec.extra ? spec.extra(items[0].dataIndex) : ''),
              footer: (items) => (spec.footer ? spec.footer(items[0].dataIndex) : spec.target ? `Objectif : ${spec.fmt(spec.target.value)}` : ''),
            },
          },
        },
      },
    });
    S.charts.push(chart);
  }

  // ---------- small building blocks ----------
  function sparkline(values) {
    const v = values.map((x) => (x == null || !isFinite(x) ? null : x));
    const pts = v.map((x, i) => [i, x]).filter((p) => p[1] != null);
    if (pts.length < 2) return '';
    const min = Math.min(...pts.map((p) => p[1])), max = Math.max(...pts.map((p) => p[1]));
    const W = 100, H = 30, span = max - min || 1, n = v.length - 1 || 1;
    const xy = pts.map(([i, x]) => [(i / n) * W, H - 3 - ((x - min) / span) * (H - 6)]);
    const d = xy.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(2)},${p[1].toFixed(2)}`).join('');
    const last = xy[xy.length - 1];
    return `<svg class="spark" viewBox="0 0 100 30" preserveAspectRatio="none" aria-hidden="true"><path d="${d}" fill="none" stroke="currentColor" stroke-width="2" vector-effect="non-scaling-stroke" stroke-linejoin="round"/><circle cx="${last[0]}" cy="${last[1]}" r="2.5" fill="currentColor" vector-effect="non-scaling-stroke"/></svg>`;
  }

  function deltaHtml(cur, prev, higherBetter, isPct) {
    if (cur == null || prev == null || !isFinite(cur) || !isFinite(prev)) return '<span class="delta flat">Pas de période précédente</span>';
    let diff, txt;
    if (isPct) { diff = (cur - prev) * 100; txt = `${diff > 0 ? '+' : ''}${nf(1).format(diff)} pt`; }
    else { if (!prev) return '<span class="delta flat">Pas de comparaison</span>'; diff = ((cur - prev) / Math.abs(prev)) * 100; txt = `${diff > 0 ? '+' : ''}${nf(0).format(diff)} %`; }
    if (Math.abs(diff) < 0.05) return '<span class="delta flat">= stable vs période préc.</span>';
    const good = higherBetter ? diff > 0 : diff < 0;
    return `<span class="delta ${good ? 'good' : 'bad'}">${icon(diff > 0 ? 'arrow-up-right' : 'arrow-down-right')}${txt}<span class="muted" style="font-weight:500">&nbsp;vs période préc.</span></span>`;
  }

  const KPI_INFO = {
    dispo: { icon: 'activity', name: 'Disponibilité des appareils', help: 'Part du temps où les appareils étaient utilisables', fmt: (v) => [v == null ? '—' : nf(2).format(v * 100), '%'], pct: true },
    pannes: { icon: 'triangle-alert', name: 'Pannes', help: 'Nombre de pannes enregistrées', fmt: (v) => [fmtN(v), v > 1 ? 'pannes' : 'panne'] },
    prev: { icon: 'calendar-check', name: 'Préventif réalisé', help: 'Part des tâches planifiées réellement faites', fmt: (v) => [v == null ? '—' : nf(1).format(v * 100), '%'], pct: true },
    cout: { icon: 'wallet', name: 'Coût de maintenance', help: 'Main-d’œuvre + pièces + sous-traitance', fmt: (v) => [fmtN(v), S.data.params.devise] },
    mtbf: { icon: 'timer', name: 'Temps moyen entre deux pannes', help: 'MTBF : plus il est long, plus l’appareil est fiable', fmt: (v) => [fmtN(v), 'h'] },
    mttr: { icon: 'wrench', name: 'Temps moyen de réparation', help: 'MTTR : durée moyenne d’une intervention de dépannage', fmt: (v) => [fmtN(v, 2), 'h'] },
    arret: { icon: 'hourglass', name: 'Temps d’arrêt total', help: 'Heures où les appareils étaient immobilisés', fmt: (v) => [fmtN(v, 1), 'h'] },
    otRetard: { icon: 'clipboard-list', name: 'Ordres de travail en retard', help: 'OT ouverts dont la date prévue est dépassée', fmt: (v) => [fmtN(v), 'OT'] },
  };
  const targetTxt = (k) => {
    const t = k.target, s = t.higher ? '≥' : '≤';
    if (k.unit === '%') return `Objectif ${s} ${fmtPct(t.obj, t.obj > 0.99 ? 1 : 0)}`;
    if (k.key === 'cout') return `Objectif ${s} ${fmtN(t.obj)} ${S.data.params.devise}`;
    return `Objectif ${s} ${fmtN(t.obj, k.key === 'mttr' ? 1 : 0)}${k.unit === 'h' ? ' h' : ''}`;
  };

  function statCard(ctx, key, sparkVals) {
    const k = ctx.kpi[key], info = KPI_INFO[key];
    const [v, unit] = info.fmt(k.value);
    const pv = ctx.prev ? ctx.prev.kpis.find((x) => x.key === key).value : null;
    return `<article class="card stat">
      <div class="stat-top"><span class="stat-ico">${icon(info.icon)}</span><div><div class="stat-name">${esc(info.name)}</div><div class="stat-help">${esc(info.help)}</div></div></div>
      <div class="stat-value"><strong>${esc(v)}</strong><span class="unit">${esc(unit)}</span></div>
      ${sparkline(sparkVals)}
      <div class="stat-foot">${kpiBadge(k)}<span>${esc(targetTxt(k))}</span></div>
      <div class="stat-foot">${deltaHtml(k.value, pv, k.target.higher, info.pct)}</div>
    </article>`;
  }

  function miniStat(ctx, key, extra) {
    const k = ctx.kpi[key], info = KPI_INFO[key];
    const [v, unit] = info.fmt(k.value);
    return `<div class="mini"><span class="k">${icon(info.icon)}${esc(info.name)}</span><span class="v">${esc(v)}<small>${esc(unit)}</small></span><span class="h">${esc(extra || info.help)}</span><span>${kpiBadge(k)}</span></div>`;
  }
  const miniPlain = (ic, name, value, unit, help) => `<div class="mini"><span class="k">${icon(ic)}${esc(name)}</span><span class="v">${value}<small>${esc(unit || '')}</small></span><span class="h">${esc(help || '')}</span></div>`;

  // ---------- pages ----------
  const PAGES = {
    apercu: { title: 'Vue d’ensemble', desc: 'L’état du parc en un coup d’œil et ce qu’il faut traiter en priorité.', render: pageOverview },
    equipements: { title: 'Équipements', desc: 'Chaque station d’anesthésie, son état et ses prochaines échéances.', render: pageAssets },
    pannes: { title: 'Pannes', desc: 'Ce qui tombe en panne, pourquoi, et combien de temps les appareils sont immobilisés.', render: pageFailures },
    preventif: { title: 'Maintenance préventive', desc: 'Les tâches planifiées : ce qui est en retard, ce qui arrive, ce qui a été fait.', render: pagePreventive },
    ot: { title: 'Ordres de travail', desc: 'Les interventions demandées et leur avancement.', render: pageWorkOrders },
    pieces: { title: 'Pièces et composants', desc: 'Les pièces remplacées, leurs références et leur coût.', render: pageParts },
    couts: { title: 'Coûts', desc: 'Où va le budget de maintenance.', render: pageCosts },
    equipe: { title: 'Équipe', desc: 'Qui intervient, sur quoi, et combien de temps.', render: pageTeam },
    guide: { title: 'Guide constructeur', desc: 'Les règles de maintenance du manuel leon plus et un lexique des indicateurs.', render: pageGuide },
    donnees: { title: 'Données et import', desc: 'Importer un fichier, télécharger le modèle et vérifier ce qui a été reconnu.', render: pageData },
  };

  function pageOverview(ctx) {
    const { res } = ctx;
    const n = ctx.assets.length;
    const counts = { good: 0, info: 0, warn: 0, bad: 0, neutral: 0 };
    ctx.assets.forEach((a) => counts[a.st.key]++);
    const ok = counts.good + counts.info;
    const urgent = ctx.actions.filter((a) => a.sev !== 'info').length;
    const showAll = S.ui.allActions;
    const list = showAll ? ctx.actions : ctx.actions.slice(0, 5);
    const months = ctx.months;
    const causes = res.causes.slice(0, 5);
    const maxCause = causes.length ? causes[0].value : 1;
    const top80 = res.causes.findIndex((c) => c.cumPct >= 0.8) + 1;
    const problem = ctx.assets.filter((a) => a.st.rank > 0).sort((a, b) => b.st.rank - a.st.rank).slice(0, 4);

    const monthlySpec = {
      id: 'ov-monthly', title: 'Pannes par mois', sub: 'Une hausse soudaine signale un appareil ou une pièce à surveiller', kind: 'bar', catLabel: 'Mois',
      labels: months.map((m) => m.label), series: [{ label: 'Pannes', data: months.map((m) => m.pannes), color: '--series-1' }],
      fmt: (v) => fmtN(v), extra: (i) => `Temps d’arrêt : ${fmtN(months[i].downtime, 1)} h`,
    };

    const html = `
      <div class="hello">
        <div><h2>Bonjour,</h2><p>Nous sommes le ${esc(longDate(ctx.ref))}. ${ok} appareil${ok > 1 ? 's' : ''} sur ${n} ${ok > 1 ? 'sont opérationnels' : 'est opérationnel'}${urgent ? `, ${plural(urgent, 'point demande', 'points demandent')} votre attention.` : ', rien d’urgent à signaler.'}</p></div>
      </div>

      <div class="row cols-4">
        ${statCard(ctx, 'dispo', months.map((m) => m.dispo))}
        ${statCard(ctx, 'pannes', months.map((m) => m.pannes))}
        ${statCard(ctx, 'prev', months.map((m) => m.prevRate))}
        ${statCard(ctx, 'cout', months.map((m) => m.cost))}
      </div>

      <div class="row cols-21">
        <section class="card" aria-labelledby="todo-t">
          <div class="card-h"><div><h2 id="todo-t">À traiter en priorité</h2><p>Classé du plus urgent au moins urgent</p></div>
            ${ctx.actions.length ? `<span class="badge ${urgent ? 'bad' : 'info'}">${icon(urgent ? 'circle-alert' : 'info')}${plural(ctx.actions.length, 'élément', 'éléments')}</span>` : ''}</div>
          <div class="card-b flush">
            ${list.length ? `<ul class="actions-list">${list.map((a) => `<li class="action"><span class="action-ico ${a.sev}">${icon(a.icon)}</span><div class="action-body"><strong>${esc(a.title)}</strong><span>${esc(a.sub)}</span></div><a class="btn sm" href="${a.href}">${esc(a.cta)}${icon('arrow-right')}</a></li>`).join('')}</ul>`
              : `<div class="all-clear"><span class="action-ico">${icon('circle-check')}</span><div><strong>Tout est à jour</strong><br><span class="muted">Aucune réparation en cours ni maintenance en retard.</span></div></div>`}
          </div>
          ${ctx.actions.length > 5 ? `<div class="card-f"><button class="btn ghost sm" type="button" id="toggleActions">${showAll ? 'Afficher moins' : `Afficher les ${ctx.actions.length} éléments`}</button></div>` : ''}
        </section>

        <section class="card" aria-labelledby="fleet-t">
          <div class="card-h"><div><h2 id="fleet-t">État du parc</h2><p>${plural(n, 'appareil suivi', 'appareils suivis')}</p></div><a class="btn ghost sm" href="#/equipements">Tout voir${icon('arrow-right')}</a></div>
          <div class="card-b">
            <div class="health-bar" role="img" aria-label="${ok} opérationnels, ${counts.warn} maintenance en retard, ${counts.bad} en réparation">
              ${[['good', ok], ['warn', counts.warn], ['bad', counts.bad]].filter(([, v]) => v).map(([k, v]) => `<span style="flex:${v};background:var(--${k})"></span>`).join('')}
            </div>
            <div class="health-legend">
              <div><span class="dot good"></span>Opérationnels<b>${ok}</b></div>
              <div><span class="dot warn"></span>Maintenance en retard<b>${counts.warn}</b></div>
              <div><span class="dot bad"></span>En réparation / en attente de pièces<b>${counts.bad}</b></div>
            </div>
            ${problem.length ? `<ul class="asset-mini">${problem.map((a) => `<li><a href="#/equipements/${encodeURIComponent(a.code)}">${esc(a.code)}</a><span class="muted">${esc(a.zone)}</span>${badge(a.st.key, a.st.label)}</li>`).join('')}</ul>` : ''}
          </div>
        </section>
      </div>

      <div class="row cols-21">
        ${chartCard(monthlySpec)}
        <section class="card" aria-labelledby="causes-t">
          <div class="card-h"><div><h2 id="causes-t">Causes les plus fréquentes</h2><p>Les 5 premières sur la période</p></div><a class="btn ghost sm" href="#/pannes">Analyser${icon('arrow-right')}</a></div>
          <div class="card-b">
            ${causes.length ? `<ol class="rank">${causes.map((c) => `<li><span class="lbl" title="${esc(c.label)}">${esc(c.label)}</span><span class="val">${fmtN(c.value)}</span><span class="bar"><i style="width:${(c.value / maxCause) * 100}%"></i></span></li>`).join('')}</ol>
              <div class="insight">${icon('lightbulb')}<span>${top80 ? `${plural(top80, 'cause explique', 'causes expliquent')} 80 % des pannes : c’est là que la prévention est la plus rentable.` : ''}</span></div>` : `<div class="empty">${icon('circle-check')}Aucune panne sur la période.</div>`}
          </div>
        </section>
      </div>

      <section class="card" aria-labelledby="rel-t">
        <div class="card-h"><div><h2 id="rel-t">Fiabilité et réactivité</h2><p>Indicateurs de suivi pour le responsable maintenance</p></div></div>
        <div class="card-b flush"><div class="mini-stats">
          ${miniStat(ctx, 'mtbf')}${miniStat(ctx, 'mttr')}${miniStat(ctx, 'arret')}${miniStat(ctx, 'otRetard')}
        </div></div>
      </section>`;
    return {
      html,
      mount() {
        mountChart(monthlySpec);
        const t = $('toggleActions');
        if (t) t.onclick = () => { S.ui.allActions = !S.ui.allActions; renderPage(); };
      },
    };
  }

  function assetCard(a) {
    const s = a.service;
    let next = '';
    if (s && s.nextAnnual != null) next = `<span>${icon('shield-check')}Maintenance annuelle</span><strong>${fmtD(s.nextAnnual)}</strong>`;
    else {
      const t = S.data.preventif.filter((x) => x.code === a.code && x.done == null && x.due != null).sort((x, y) => x.due - y.due)[0];
      next = t ? `<span>${icon('calendar-check')}Prochaine tâche</span><strong>${fmtD(t.due)}</strong>` : `<span>${icon('calendar-check')}Aucune tâche planifiée</span>`;
    }
    return `<button type="button" class="asset" data-code="${esc(a.code)}">
      <div class="asset-head"><span class="asset-ico">${icon('monitor-smartphone')}</span>
        <div class="asset-title"><strong>${esc(a.code)} · ${esc(a.model || a.name)}</strong><span>${esc([a.zone, a.line].filter(Boolean).join(' · '))}</span>${badge(a.st.key, a.st.label)}</div></div>
      <div class="asset-kv">
        <div><span>Disponibilité</span><b>${fmtPct(a.dispo, 1)}</b></div>
        <div><span>Pannes</span><b>${fmtN(a.pannes)}</b></div>
        <div><span>Criticité</span><b style="font-size:13px">${critHtml(a.crit)}</b></div>
      </div>
      <div class="asset-next">${next}</div>
    </button>`;
  }

  function pageAssets(ctx) {
    const u = S.ui.assets || (S.ui.assets = { status: 'all', crit: '', view: 'grid', q: '' });
    const groups = { all: ctx.assets, good: ctx.assets.filter((a) => a.st.key === 'good' || a.st.key === 'info'), warn: ctx.assets.filter((a) => a.st.key === 'warn'), bad: ctx.assets.filter((a) => a.st.key === 'bad') };
    const q = norm(u.q);
    const list = groups[u.status].filter((a) => (!u.crit || a.crit === u.crit) && (!q || norm(`${a.code} ${a.model} ${a.name} ${a.zone} ${a.line} ${a.serial}`).includes(q)));
    const crits = C.CRIT_ORDER.filter((c) => ctx.assets.some((a) => a.crit === c));
    const tab = (k, label) => `<button type="button" role="tab" aria-selected="${u.status === k}" data-st="${k}">${esc(label)}<span class="count">${groups[k].length}</span></button>`;
    const cols = [
      { label: 'Équipement', html: (a) => eqCell(a.code), sortVal: (a) => a.code },
      { label: 'Service / salle', html: (a) => `${esc(a.zone)}<span class="sub">${esc(a.line)}</span>`, sortVal: (a) => a.zone },
      { label: 'État', html: (a) => badge(a.st.key, a.st.label), sortVal: (a) => a.st.rank },
      { label: 'Criticité', html: (a) => critHtml(a.crit), sortVal: (a) => CRIT_LEVEL[a.crit] || 0 },
      { label: 'N° de série', html: (a) => esc(a.serial || '—'), nw: true },
      { label: 'Compteur', n: true, html: (a) => (a.hours != null ? `${fmtN(a.hours)} h` : '—'), sortVal: (a) => a.hours },
      { label: 'Disponibilité', n: true, html: (a) => fmtPct(a.dispo, 2), sortVal: (a) => a.dispo },
      { label: 'Pannes', n: true, html: (a) => fmtN(a.pannes), sortVal: (a) => a.pannes },
      { label: 'Coût', n: true, html: (a) => fmtMoney(a.cost), sortVal: (a) => a.cost },
      { label: 'Maintenance annuelle', nw: true, html: (a) => (a.service && a.service.nextAnnual != null ? fmtD(a.service.nextAnnual) : '—'), sortVal: (a) => (a.service ? a.service.nextAnnual : null) },
    ];
    const html = `
      <div class="toolbar">
        <div class="left">
          <div class="tabs" role="tablist" aria-label="Filtrer par état">${tab('all', 'Tous')}${tab('good', 'Opérationnels')}${tab('warn', 'Maintenance en retard')}${tab('bad', 'En réparation')}</div>
        </div>
        <div class="right">
          ${searchBox('assetSearch', 'Rechercher un appareil, une salle…')}
          <select id="assetCrit" aria-label="Criticité"><option value="">Toutes criticités</option>${crits.map((c) => `<option ${u.crit === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
          <div class="view-toggle" role="group" aria-label="Affichage">
            <button type="button" data-v="grid" aria-pressed="${u.view === 'grid'}">${icon('layout-grid')}Cartes</button>
            <button type="button" data-v="table" aria-pressed="${u.view === 'table'}">${icon('table')}Tableau</button>
          </div>
        </div>
      </div>
      <div id="assetList">${
        !list.length ? `<div class="card"><div class="empty">${icon('search-x')}Aucun appareil ne correspond à ces critères.</div></div>`
        : u.view === 'grid' ? `<div class="asset-grid">${list.map(assetCard).join('')}</div>`
        : `<section class="card">${dataTable({ id: 'assetTable', caption: 'Liste des équipements', cols, rows: list, href: (a) => `#/equipements/${encodeURIComponent(a.code)}` })}</section>`
      }</div>`;
    return {
      html,
      mount() {
        const s = $('assetSearch');
        s.value = u.q;
        s.addEventListener('input', () => { u.q = s.value; renderPage(true); });
        document.querySelectorAll('[data-st]').forEach((b) => (b.onclick = () => { u.status = b.dataset.st; renderPage(true); }));
        $('assetCrit').onchange = (e) => { u.crit = e.target.value; renderPage(true); };
        document.querySelectorAll('[data-v]').forEach((b) => (b.onclick = () => { u.view = b.dataset.v; renderPage(true); }));
        document.querySelectorAll('.asset[data-code]').forEach((b) => (b.onclick = () => (location.hash = `#/equipements/${encodeURIComponent(b.dataset.code)}`)));
        wireTable('assetTable');
      },
    };
  }

  function pageFailures(ctx) {
    const { res } = ctx;
    const causes = res.causes.slice(0, 10);
    const fam = res.families.slice(0, 10);
    const n80 = res.causes.findIndex((c) => c.cumPct >= 0.8) + 1;
    const dispoT = ctx.kpi.dispo.target.obj;
    const worstMonth = ctx.months.filter((m) => m.dispo != null).sort((a, b) => a.dispo - b.dispo)[0];
    const specs = [
      { id: 'f-causes', title: 'Pourquoi les appareils tombent en panne', sub: 'Causes des pannes correctives · 10 premières', kind: 'hbar', tall: true, labels: causes.map((c) => c.label), series: [{ label: 'Pannes', data: causes.map((c) => c.value), color: '--series-1' }], fmt: (v) => fmtN(v), extra: (i) => `Cumul : ${fmtPct(causes[i].cumPct, 0)} · arrêt ${fmtN(causes[i].downtime, 1)} h`, extraCols: { head: ['Cumul', 'Arrêt (h)'], row: (i) => [fmtPct(causes[i].cumPct, 0), fmtN(causes[i].downtime, 1)] }, insight: causes.length ? `<strong>${esc(causes[0].label)}</strong> est la première cause (${plural(causes[0].value, 'panne', 'pannes')}). ${plural(n80, 'cause regroupe', 'causes regroupent')} 80 % des pannes.` : '' },
      { id: 'f-fam', title: 'Quels modules sont concernés', sub: 'Pannes par sous-ensemble de l’appareil', kind: 'hbar', tall: true, labels: fam.map((c) => c.label), series: [{ label: 'Pannes', data: fam.map((c) => c.value), color: '--series-1' }], fmt: (v) => fmtN(v), insight: fam.length ? `Le module <strong>${esc(fam[0].label)}</strong> concentre le plus de pannes : vérifier le stock de pièces associées.` : '' },
      { id: 'f-dispo', title: 'Disponibilité mois par mois', sub: `Ligne pointillée : objectif ${fmtPct(dispoT, 1)}`, kind: 'line', pct: true, catLabel: 'Mois', labels: ctx.months.map((m) => m.label), series: [{ label: 'Disponibilité', data: ctx.months.map((m) => m.dispo), color: '--series-1' }], target: { label: 'Objectif', value: dispoT }, fmt: (v) => fmtPct(v, 2), insight: worstMonth ? `Mois le plus difficile : <strong>${esc(worstMonth.label)}</strong> (${fmtPct(worstMonth.dispo, 2)}, ${plural(worstMonth.pannes, 'panne', 'pannes')}).` : '' },
      { id: 'f-crit', title: 'Pannes selon la criticité des appareils', sub: 'Les appareils critiques doivent tomber le moins possible en panne', kind: 'bar', catLabel: 'Criticité', labels: res.critData.map((c) => c.label), series: [{ label: 'Pannes', data: res.critData.map((c) => c.pannes), color: '--series-1' }], fmt: (v) => fmtN(v), extra: (i) => `Arrêt : ${fmtN(res.critData[i].downtime, 1)} h`, extraCols: { head: ['Arrêt (h)'], row: (i) => [fmtN(res.critData[i].downtime, 1)] } },
    ];
    const P = res.recentPannes;
    const hasAlarm = P.some((p) => p.alarm), hasPart = P.some((p) => p.part);
    const cols = [
      { label: 'Date', nw: true, html: (p) => fmtD(p.date), sortVal: (p) => p.date },
      { label: 'Équipement', html: (p) => eqCell(p.code), sortVal: (p) => p.code },
      { label: 'Problème', wide: true, html: (p) => `<strong>${esc(p.desc || '—')}</strong><span class="sub">${esc(p.cause)}</span>`, sortVal: (p) => p.desc },
      { label: 'Module', html: (p) => esc(p.family || '—'), sortVal: (p) => p.family },
    ];
    if (hasAlarm) cols.push({ label: 'Code alarme', html: (p) => (p.alarm ? `<span class="tag">${esc(p.alarm)}</span>` : '<span class="muted">—</span>'), sortVal: (p) => p.alarm });
    if (hasPart) cols.push({ label: 'Réf. composant', html: (p) => (p.part ? `<span class="tag">${esc(p.part)}</span>` : '<span class="muted">—</span>'), sortVal: (p) => p.part });
    cols.push(
      { label: 'Arrêt', n: true, html: (p) => `${fmtN(p.downtime, 1)} h`, sortVal: (p) => p.downtime },
      { label: 'Réparation', n: true, html: (p) => `${fmtN(p.repair, 2)} h`, sortVal: (p) => p.repair },
      { label: 'Intervenant', html: (p) => `${esc(p.resp || '—')}${p.include ? '' : '<span class="sub">exclue des KPI</span>'}`, sortVal: (p) => p.resp }
    );
    const html = `
      <section class="card"><div class="card-b flush"><div class="mini-stats">
        ${miniStat(ctx, 'pannes', ctx.kpi.pannes.detail)}
        ${miniPlain('hourglass', 'Arrêt dû aux pannes', fmtN(res.fleet.reduce((a, e) => a + e.downtime, 0), 1), 'h', 'Heures d’immobilisation sur pannes')}
        ${miniStat(ctx, 'mttr')}${miniStat(ctx, 'mtbf')}${miniStat(ctx, 'dispo')}
      </div></div></section>
      <div class="row cols-2">${chartCard(specs[0])}${chartCard(specs[1])}</div>
      <div class="row cols-2">${chartCard(specs[2])}${chartCard(specs[3])}</div>
      <section class="card" aria-labelledby="log-t">
        <div class="card-h"><div><h2 id="log-t">Registre des pannes</h2><p>Cliquez sur un en-tête pour trier, ou sur une ligne pour ouvrir l’appareil</p></div><div class="tools">${searchBox('pannesSearch', 'Rechercher une panne, une cause…')}</div></div>
        <div class="card-b flush">${dataTable({ id: 'pannesTable', caption: 'Registre des pannes', cols, rows: P, empty: 'Aucune panne sur la période.', href: (p) => `#/equipements/${encodeURIComponent(p.code)}` })}</div>
      </section>`;
    return { html, mount() { specs.forEach(mountChart); wireTable('pannesTable', $('pannesSearch')); } };
  }

  function pagePreventive(ctx) {
    const { res } = ctx;
    const u = S.ui.prev || (S.ui.prev = { tab: res.prevLate.length ? 'late' : 'soon' });
    const k = ctx.kpi.prev;
    const all = S.data.preventif.filter((t) => t.due != null && t.due >= res.from && t.due <= Math.max(res.to, ctx.ref + 30) && scopeOK(t.code, t.zone)).sort((a, b) => b.due - a.due);
    const doneInP = all.filter((t) => t.done != null && t.due <= res.toEff).length;
    const rows = { late: res.prevLate, soon: res.prevSoon, all };
    const statusBadge = (t) => (t.done != null ? badge('good', `Réalisée le ${fmtD(t.done)}`) : t.due < ctx.ref ? badge('bad', `En retard de ${plural(ctx.ref - t.due, 'jour', 'jours')}`) : t.due === ctx.ref ? badge('warn', 'Aujourd’hui') : badge('info', `Dans ${plural(t.due - ctx.ref, 'jour', 'jours')}`));
    const cols = [
      { label: 'Échéance', nw: true, html: (t) => fmtD(t.due), sortVal: (t) => t.due },
      { label: 'Équipement', html: (t) => eqCell(t.code), sortVal: (t) => t.code },
      { label: 'Opération', wide: true, html: (t) => `<strong>${esc(t.op)}</strong>${t.id ? `<span class="sub">Gamme ${esc(t.id)}</span>` : ''}`, sortVal: (t) => t.op },
      { label: 'Fréquence', html: (t) => esc(t.freq || '—'), sortVal: (t) => t.freq },
      { label: 'Responsable', html: (t) => esc(t.resp || '—'), sortVal: (t) => t.resp },
      { label: 'Statut', html: statusBadge, sortVal: (t) => (t.done != null ? 2 : t.due < ctx.ref ? 0 : 1) },
    ];
    const tab = (key, label) => `<button type="button" role="tab" aria-selected="${u.tab === key}" data-tab="${key}">${esc(label)}<span class="count">${rows[key].length}</span></button>`;
    const spec = {
      id: 'p-month', title: 'Tâches prévues et réalisées par mois', sub: 'Quand les deux barres ont la même hauteur, le plan est tenu', kind: 'bar', grouped: true, catLabel: 'Mois',
      labels: ctx.months.map((m) => m.label),
      series: [{ label: 'Prévues', data: ctx.months.map((m) => m.due), color: '--series-1' }, { label: 'Réalisées', data: ctx.months.map((m) => m.done), color: '--series-3' }],
      fmt: (v) => fmtN(v), footer: (i) => `Taux : ${fmtPct(ctx.months[i].prevRate, 0)}`,
    };
    const html = `
      <section class="card"><div class="card-b flush"><div class="mini-stats">
        <div class="mini"><span class="k">${icon('calendar-check')}Taux de réalisation</span><span class="v">${k.value == null ? '—' : nf(1).format(k.value * 100)}<small>%</small></span><div class="progress" role="img" aria-label="${fmtPct(k.value, 0)} réalisé"><i style="width:${(k.value || 0) * 100}%"></i></div><span class="h">${esc(k.detail)}</span><span>${kpiBadge(k)}</span></div>
        ${miniPlain('calendar-x', 'En retard', `<span style="color:var(--bad-fg)">${fmtN(res.prevLate.length)}</span>`, 'tâches', 'Échéance dépassée, non réalisée')}
        ${miniPlain('calendar-clock', 'À faire sous 30 jours', fmtN(res.prevSoon.length), 'tâches', 'À planifier dès maintenant')}
        ${miniPlain('list-checks', 'Réalisées sur la période', fmtN(doneInP), 'tâches', 'Tâches faites dont l’échéance est dans la période')}
      </div></div></section>
      ${chartCard(spec)}
      <section class="card" aria-labelledby="plan-t">
        <div class="card-h"><div><h2 id="plan-t">Plan de maintenance</h2><p>Les tâches en retard sont à traiter en premier</p></div>
          <div class="tools"><div class="tabs" role="tablist" aria-label="Filtrer les tâches">${tab('late', 'En retard')}${tab('soon', 'Sous 30 jours')}${tab('all', 'Tout le plan')}</div>${searchBox('prevSearch', 'Rechercher…')}</div></div>
        <div class="card-b flush">${dataTable({ id: 'prevTable', caption: 'Plan de maintenance préventive', cols, rows: rows[u.tab], empty: u.tab === 'late' ? 'Aucune tâche en retard. Bravo !' : 'Aucune tâche.', href: (t) => `#/equipements/${encodeURIComponent(t.code)}` })}</div>
      </section>`;
    return {
      html,
      mount() {
        mountChart(spec);
        wireTable('prevTable', $('prevSearch'));
        document.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { u.tab = b.dataset.tab; renderPage(true); }));
      },
    };
  }

  function pageWorkOrders(ctx) {
    const { res } = ctx;
    const u = S.ui.ot || (S.ui.ot = { tab: 'open' });
    const scope = S.data.ot.filter((o) => scopeOK(o.code, o.zone));
    const open = scope.filter((o) => !isClosedOT(o));
    const late = open.filter((o) => o.planned != null && o.planned < ctx.ref);
    const urgent = open.filter((o) => /urgent/i.test(o.prio));
    const closedP = scope.filter((o) => norm(o.status) === 'cloture' && o.closed != null && o.closed >= res.from && o.closed <= res.to);
    const delays = closedP.filter((o) => o.created != null).map((o) => o.closed - o.created);
    const avgDelay = delays.length ? delays.reduce((a, b) => a + b, 0) / delays.length : null;
    const byStatus = {};
    open.forEach((o) => (byStatus[o.status || 'Ouvert'] = (byStatus[o.status || 'Ouvert'] || 0) + 1));
    const prioBadge = (p) => (/urgent/i.test(p) ? badge('bad', p) : /haute/i.test(p) ? badge('warn', p) : p ? badge('neutral', p) : '—');
    const stBadge = (o) => (isClosedOT(o) ? badge('good', o.status) : o.planned != null && o.planned < ctx.ref ? badge('bad', `${o.status || 'Ouvert'} · en retard`) : /attente/i.test(o.status) ? badge('warn', o.status) : badge('info', o.status || 'Ouvert'));
    const rows = { open: open.sort((a, b) => (a.planned || 0) - (b.planned || 0)), all: scope.slice().sort((a, b) => (b.created || 0) - (a.created || 0)) };
    const cols = [
      { label: 'N° OT', nw: true, html: (o) => `<strong>${esc(o.id)}</strong>`, sortVal: (o) => o.id },
      { label: 'Équipement', html: (o) => eqCell(o.code), sortVal: (o) => o.code },
      { label: 'Travail demandé', wide: true, html: (o) => `<strong>${esc(o.desc)}</strong><span class="sub">${esc(o.type)}</span>`, sortVal: (o) => o.desc },
      { label: 'Priorité', html: (o) => prioBadge(o.prio), sortVal: (o) => ({ urgente: 0, haute: 1, normale: 2, basse: 3 }[norm(o.prio)] ?? 4) },
      { label: 'Responsable', html: (o) => esc(o.resp || '—'), sortVal: (o) => o.resp },
      { label: 'Créé le', nw: true, html: (o) => fmtD(o.created), sortVal: (o) => o.created },
      { label: 'Prévu le', nw: true, html: (o) => fmtD(o.planned), sortVal: (o) => o.planned },
      { label: 'Statut', html: stBadge, sortVal: (o) => o.status },
    ];
    const col = (ic, k, v, h, tone) => `<div class="board-col"><span class="k">${icon(ic)}${esc(k)}</span><span class="v"${tone ? ` style="color:var(--${tone}-fg)"` : ''}>${v}</span><span class="h">${esc(h)}</span></div>`;
    const html = `
      <div class="board">
        ${col('folder-open', 'Ouverts', fmtN(open.length), Object.entries(byStatus).map(([k, v]) => `${v} ${k.toLowerCase()}`).join(' · ') || 'Aucun')}
        ${col('clock-alert', 'En retard', fmtN(late.length), 'Date prévue dépassée', late.length ? 'bad' : '')}
        ${col('siren', 'Urgents ouverts', fmtN(urgent.length), 'Priorité « Urgente »', urgent.length ? 'warn' : '')}
        ${col('circle-check-big', 'Clôturés sur la période', fmtN(closedP.length), 'Travaux terminés')}
        ${col('timer', 'Délai moyen de clôture', avgDelay == null ? '—' : nf(1).format(avgDelay), 'jours entre création et clôture')}
      </div>
      <section class="card" aria-labelledby="ot-t">
        <div class="card-h"><div><h2 id="ot-t">Ordres de travail</h2><p>Triés par date prévue : le plus urgent en haut</p></div>
          <div class="tools"><div class="tabs" role="tablist" aria-label="Filtrer les OT"><button type="button" role="tab" data-tab="open" aria-selected="${u.tab === 'open'}">Ouverts<span class="count">${open.length}</span></button><button type="button" role="tab" data-tab="all" aria-selected="${u.tab === 'all'}">Tous<span class="count">${scope.length}</span></button></div>${searchBox('otSearch', 'Rechercher un OT…')}</div></div>
        <div class="card-b flush">${dataTable({ id: 'otTable', caption: 'Ordres de travail', cols, rows: rows[u.tab], empty: 'Aucun ordre de travail.', href: (o) => `#/equipements/${encodeURIComponent(o.code)}` })}</div>
      </section>`;
    return { html, mount() { wireTable('otTable', $('otSearch')); document.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { u.tab = b.dataset.tab; renderPage(true); })); } };
  }

  function pageParts(ctx) {
    const parts = ctx.res.parts;
    const qty = parts.reduce((a, p) => a + p.qty, 0);
    const cost = parts.reduce((a, p) => a + p.cost, 0);
    const top = parts.slice(0, 10);
    const spec = {
      id: 'parts-top', title: 'Pièces les plus remplacées', sub: 'Nombre de remplacements sur la période · utile pour dimensionner le stock', kind: 'hbar', tall: true,
      labels: top.map((p) => p.label || p.ref), series: [{ label: 'Remplacements', data: top.map((p) => p.qty), color: '--series-1' }],
      fmt: (v) => fmtN(v), extra: (i) => `${top[i].ref ? `Réf. ${top[i].ref} · ` : ''}coût ${fmtMoney(top[i].cost)}`, extraCols: { head: ['Coût'], row: (i) => [fmtMoney(top[i].cost)] },
    };
    const byCost = parts.slice().sort((a, b) => b.cost - a.cost)[0];
    const cols = [
      { label: 'Référence', html: (p) => (p.ref ? `<span class="tag">${esc(p.ref)}</span>` : '<span class="muted">—</span>'), sortVal: (p) => p.ref },
      { label: 'Pièce / composant', html: (p) => `<strong>${esc(p.label || '—')}</strong>`, sortVal: (p) => p.label },
      { label: 'Remplacements', n: true, html: (p) => fmtN(p.qty), sortVal: (p) => p.qty },
      { label: 'Coût des pièces', n: true, html: (p) => fmtMoney(p.cost), sortVal: (p) => p.cost },
      { label: 'Appareils concernés', html: (p) => p.eqs.map((c) => `<a href="#/equipements/${encodeURIComponent(c)}">${esc(c)}</a>`).join(', '), sortVal: (p) => p.eqs.length },
      { label: 'Dernier remplacement', nw: true, html: (p) => fmtD(p.last), sortVal: (p) => p.last },
    ];
    const html = `
      <section class="card"><div class="card-b flush"><div class="mini-stats">
        ${miniPlain('package', 'Pièces remplacées', fmtN(qty), '', 'Nombre total de remplacements')}
        ${miniPlain('wallet', 'Coût des pièces', fmtN(cost), S.data.params.devise, 'Sur la période')}
        ${miniPlain('tags', 'Références différentes', fmtN(parts.length), '', 'Pièces distinctes utilisées')}
        ${miniPlain('coins', 'Poste le plus coûteux', byCost ? esc(trunc(byCost.label || byCost.ref, 22)) : '—', '', byCost ? fmtMoney(byCost.cost) : '')}
      </div></div></section>
      ${chartCard(spec)}
      <section class="card" aria-labelledby="parts-t">
        <div class="card-h"><div><h2 id="parts-t">Catalogue des pièces utilisées</h2><p>Regroupées par référence${S.isDemo ? ' · références DEMO-REF à remplacer par la liste officielle Löwenstein' : ''}</p></div><div class="tools">${searchBox('partsSearch', 'Rechercher une pièce, une référence…')}</div></div>
        <div class="card-b flush">${dataTable({ id: 'partsTable', caption: 'Pièces remplacées', cols, rows: parts, empty: 'Aucune pièce saisie sur la période.' })}</div>
      </section>`;
    return { html, mount() { mountChart(spec); wireTable('partsTable', $('partsSearch')); } };
  }

  function pageCosts(ctx) {
    const { res } = ctx;
    const m = ctx.months;
    const tot = { mo: 0, parts: 0, other: 0, total: 0 };
    m.forEach((x) => { tot.mo += x.mo; tot.parts += x.parts; tot.other += x.other; tot.total += x.cost; });
    const inP = (s) => s != null && s >= res.from && s <= res.to;
    const interv = S.data.interventions.filter((i) => inP(i.date) && scopeOK(i.code, i.zone) && norm(i.status) !== 'annulee');
    const byType = {};
    interv.forEach((i) => { const t = C.typeGroup(i.type); byType[t] = (byType[t] || 0) + (i.costTotal || 0); });
    const types = Object.entries(byType).sort((a, b) => b[1] - a[1]);
    const maxT = types.length ? types[0][1] : 1;
    const corrCost = byType.Correctif || 0;
    const perPanne = ctx.kpi.pannes.value ? corrCost / ctx.kpi.pannes.value : null;
    const dev = S.data.params.devise;
    const costs = res.costByEq.slice(0, 10);
    const series3 = (src) => [
      { label: 'Main-d’œuvre', data: src.map((x) => x.mo), color: '--series-1' },
      { label: 'Pièces', data: src.map((x) => x.parts), color: '--series-2' },
      { label: 'Autres (sous-traitance…)', data: src.map((x) => x.other), color: '--series-3' },
    ].filter((s) => s.data.some((v) => v > 0));
    const specs = [
      { id: 'c-month', title: 'Dépenses mois par mois', sub: `En ${dev}, par nature de coût`, kind: 'bar', stacked: true, catLabel: 'Mois', labels: m.map((x) => x.label), series: series3(m), fmt: (v) => `${fmtN(v)} ${dev}`, footer: (i) => `Total : ${fmtN(m[i].cost)} ${dev}` },
      { id: 'c-eq', title: 'Coût par appareil', sub: `10 appareils les plus coûteux, en ${dev}`, kind: 'hbar', stacked: true, tall: true, labels: costs.map((c) => c.label), series: series3(costs), fmt: (v) => `${fmtN(v)} ${dev}`, footer: (i) => `Total : ${fmtN(costs[i].total)} ${dev}`, insight: costs.length ? `<strong>${esc(costs[0].label)}</strong> coûte le plus cher (${fmtN(costs[0].total)} ${esc(dev)}), dont ${fmtPct(costs[0].parts / (costs[0].total || 1), 0)} en pièces.` : '' },
    ];
    const html = `
      <section class="card"><div class="card-b flush"><div class="mini-stats">
        ${miniStat(ctx, 'cout', 'Hors interventions annulées')}
        ${miniPlain('hard-hat', 'Main-d’œuvre', fmtN(tot.mo), dev, fmtPct(tot.mo / (tot.total || 1), 0) + ' du total')}
        ${miniPlain('package', 'Pièces', fmtN(tot.parts), dev, fmtPct(tot.parts / (tot.total || 1), 0) + ' du total')}
        ${miniPlain('handshake', 'Sous-traitance et autres', fmtN(tot.other), dev, fmtPct(tot.other / (tot.total || 1), 0) + ' du total')}
        ${miniPlain('receipt', 'Coût moyen d’une panne', fmtN(perPanne), dev, 'Coût correctif ÷ nombre de pannes')}
      </div></div></section>
      <div class="row cols-21">${chartCard(specs[0])}
        <section class="card" aria-labelledby="ctype-t"><div class="card-h"><div><h2 id="ctype-t">Par type de maintenance</h2><p>Le préventif évite des pannes plus coûteuses</p></div></div>
          <div class="card-b">${types.length ? `<ol class="rank">${types.map(([t, v]) => `<li><span class="lbl">${esc(t)}</span><span class="val">${fmtN(v)} ${esc(dev)}</span><span class="bar"><i style="width:${(v / maxT) * 100}%"></i></span></li>`).join('')}</ol>` : `<div class="empty">Aucune dépense.</div>`}</div></section>
      </div>
      ${chartCard(specs[1])}`;
    return { html, mount() { specs.forEach(mountChart); } };
  }

  function pageTeam(ctx) {
    const { res } = ctx;
    const resp = res.byResp.slice(0, 12);
    const types = res.respTypes.filter((t) => resp.some((r) => r.hours[t] > 0));
    const colors = ['--series-1', '--series-2', '--series-3', '--series-4'];
    const spec = {
      id: 't-load', title: 'Heures d’intervention par personne', sub: 'Réparties par type de maintenance', kind: 'hbar', stacked: true, tall: true,
      labels: resp.map((r) => r.label), series: types.map((t) => ({ label: t, data: resp.map((r) => r.hours[t]), color: colors[res.respTypes.indexOf(t)] })),
      fmt: (v) => `${fmtN(v, 1)} h`, footer: (i) => `Total : ${fmtN(resp[i].total, 1)} h · ${plural(resp[i].count, 'intervention', 'interventions')}`,
    };
    const inP = (s) => s != null && s >= res.from && s <= res.to;
    const pannes = S.data.pannes.filter((p) => inP(p.date) && scopeOK(p.code, p.zone));
    const rows = res.byResp.map((r) => {
      const mine = pannes.filter((p) => (p.resp || 'Non renseigné') === r.label);
      const rep = mine.filter((p) => p.repair != null);
      const prevDone = S.data.preventif.filter((t) => (t.resp || 'Non renseigné') === r.label && t.done != null && inP(t.done) && scopeOK(t.code, t.zone)).length;
      return Object.assign({}, r, { pannes: mine.length, mttr: rep.length ? rep.reduce((a, p) => a + p.repair, 0) / rep.length : null, prevDone });
    });
    const cols = [
      { label: 'Intervenant', html: (r) => `<strong>${esc(r.label)}</strong>`, sortVal: (r) => r.label },
      { label: 'Interventions', n: true, html: (r) => fmtN(r.count), sortVal: (r) => r.count },
      { label: 'Heures', n: true, html: (r) => `${fmtN(r.total, 1)} h`, sortVal: (r) => r.total },
      { label: 'Pannes traitées', n: true, html: (r) => fmtN(r.pannes), sortVal: (r) => r.pannes },
      { label: 'Réparation moyenne', n: true, html: (r) => (r.mttr == null ? '—' : `${fmtN(r.mttr, 2)} h`), sortVal: (r) => r.mttr },
      { label: 'Tâches préventives faites', n: true, html: (r) => fmtN(r.prevDone), sortVal: (r) => r.prevDone },
    ];
    const html = `
      ${chartCard(spec)}
      <section class="card" aria-labelledby="team-t"><div class="card-h"><div><h2 id="team-t">Activité par intervenant</h2><p>Sur la période sélectionnée</p></div></div>
        <div class="card-b flush">${dataTable({ id: 'teamTable', caption: 'Activité par intervenant', cols, rows, empty: 'Aucune intervention sur la période.' })}</div></section>`;
    return { html, mount() { mountChart(spec); wireTable('teamTable'); } };
  }

  function pageGuide() {
    const R = C.LEON_REFERENCE;
    const icons = ['calendar-check', 'calendar-range', 'calendar-cog', 'gauge', 'flask-conical'];
    const consCols = [
      { label: 'Pièce / consommable', html: (r) => `<strong>${esc(r[0])}</strong>`, sortVal: (r) => r[0] },
      { label: 'Quand la remplacer', html: (r) => esc(r[1]), sort: false },
    ];
    const gloss = [
      ['Disponibilité', 'Part du temps où l’appareil pouvait être utilisé. 99 % sur un mois ≈ 1 heure d’arrêt pour 100 heures d’ouverture.'],
      ['MTBF (temps moyen entre pannes)', 'Durée moyenne de bon fonctionnement entre deux pannes. Plus il est élevé, plus l’appareil est fiable.'],
      ['MTTR (temps moyen de réparation)', 'Durée moyenne d’un dépannage, de la prise en charge à la remise en service.'],
      ['Maintenance préventive', 'Tâches planifiées à l’avance (remplacement de consommables, contrôles) pour éviter les pannes.'],
      ['Maintenance corrective', 'Intervention après une panne pour remettre l’appareil en état.'],
      ['CST', 'Contrôle de sécurité technique : vérification réglementaire annuelle, consignée dans le registre du dispositif médical.'],
      ['Ordre de travail (OT)', 'Demande d’intervention : quoi faire, sur quel appareil, par qui et pour quand.'],
      ['Criticité', 'Importance de l’appareil pour l’activité : un appareil « Critique » doit être réparé en priorité.'],
    ];
    const html = `
      <div class="notice">${icon('book-open')}<div>Source : ${esc(R.source)}. Les opérations marquées « Technicien agréé » ne peuvent être réalisées que par du personnel formé et habilité par Löwenstein Medical, avec des pièces d’origine.</div></div>
      <div class="row cols-12">
        <section class="card" aria-labelledby="int-t"><div class="card-h"><div><h2 id="int-t">Calendrier de maintenance</h2><p>Intervalles imposés par le constructeur</p></div></div>
          <div class="card-b"><ol class="timeline">${R.intervals.map((r, i) => `<li><span class="t-ico">${icon(icons[i] || 'calendar')}</span><div><h3>${esc(r[0])}</h3><p>${esc(r[1])}</p><span class="who">${icon('user-check')}${esc(r[2])}</span></div></li>`).join('')}</ol></div></section>
        <section class="card" aria-labelledby="cons-t"><div class="card-h"><div><h2 id="cons-t">Consommables et pièces d’usure</h2><p>Durée de vie indiquée par le manuel</p></div></div>
          <div class="card-b flush">${dataTable({ id: 'consTable', caption: 'Durée de vie des consommables', cols: consCols, rows: R.consumables })}</div></section>
      </div>
      <section class="card" aria-labelledby="gl-t"><div class="card-h"><div><h2 id="gl-t">Lexique</h2><p>Les termes utilisés dans ce tableau de bord, en clair</p></div></div>
        <div class="card-b"><dl class="glossary">${gloss.map(([t, d]) => `<div class="gloss"><dt>${esc(t)}</dt><dd>${esc(d)}</dd></div>`).join('')}</dl></div></section>`;
    return { html, mount() { wireTable('consTable'); } };
  }

  function pageData(ctx) {
    const rep = S.report;
    const names = { equipements: 'Liste des équipements', pannes: 'Registre des pannes', interventions: 'Historique des interventions', preventif: 'Plan préventif', ot: 'Ordres de travail' };
    const P = S.data.params;
    const checks = rep
      ? Object.entries(rep.tables).map(([k, t]) => (t
        ? `<li>${icon('circle-check', 'ok')}<div><strong>${names[k]}</strong> : feuille « ${esc(t.sheet)} », ${plural(t.rows, 'ligne', 'lignes')}${t.missing.length ? `<br><span class="muted">Colonnes facultatives absentes : ${esc(t.missing.join(', '))}</span>` : ''}</div></li>`
        : `<li>${icon('circle-alert', 'ko')}<div><strong>${names[k]}</strong> : non trouvé dans le fichier</div></li>`)).join('')
      : `<li>${icon('info', 'ok')}<div>Données de démonstration générées dans le navigateur (8 stations LM leon fictives).</div></li>`;
    const tcols = [
      { label: 'Indicateur', html: (k) => `<strong>${esc(KPI_INFO[k.key].name)}</strong>`, sortVal: (k) => k.label },
      { label: 'Objectif', n: true, html: (k) => esc(targetTxt(k).replace('Objectif ', '')), sort: false },
      { label: 'Seuil d’alerte', n: true, html: (k) => esc(k.unit === '%' ? fmtPct(k.target.thr, 1) : fmtN(k.target.thr, k.key === 'mttr' ? 1 : 0)), sort: false },
      { label: 'Valeur actuelle', n: true, html: (k) => esc(KPI_INFO[k.key].fmt(k.value).join(' ')), sort: false },
      { label: 'État', html: (k) => kpiBadge(k), sort: false },
    ];
    const html = `
      <div class="row cols-2">
        <section class="card" aria-labelledby="imp-t"><div class="card-h"><div><h2 id="imp-t">Importer des données</h2><p>Le fichier est lu sur cet ordinateur : rien n’est envoyé sur Internet</p></div></div>
          <div class="card-b">
            <div class="dropzone">${icon('file-spreadsheet')}<strong>Glissez votre fichier Excel ici</strong><p>Formats acceptés : modèle SAMQ « Dashboard Maintenance Industrielle » ou modèle GMAO leon. Les colonnes sont reconnues par leur nom.</p>
              <label class="btn primary">${icon('upload')}Choisir un fichier<input type="file" id="file2" accept=".xlsx,.xlsm,.xls,.ods"></label></div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:16px">
              <button class="btn" type="button" id="dlTemplate">${icon('download')}Modèle Excel vierge</button>
              <button class="btn" type="button" id="dlExport">${icon('file-down')}Exporter les données</button>
              <button class="btn" type="button" id="dlDemo">${icon('rotate-ccw')}Recharger la démo</button>
            </div>
          </div></section>
        <section class="card" aria-labelledby="src-t"><div class="card-h"><div><h2 id="src-t">Source actuelle</h2><p>${esc(S.source)}${S.isDemo ? ' · données fictives' : ''}</p></div></div>
          <div class="card-b">
            <dl class="kv-list">
              <dt>Équipements</dt><dd>${fmtN(S.data.equipements.length)}</dd>
              <dt>Pannes</dt><dd>${fmtN(S.data.pannes.length)}</dd>
              <dt>Interventions</dt><dd>${fmtN(S.data.interventions.length)}</dd>
              <dt>Tâches préventives</dt><dd>${fmtN(S.data.preventif.length)}</dd>
              <dt>Ordres de travail</dt><dd>${fmtN(S.data.ot.length)}</dd>
              <dt>Devise</dt><dd>${esc(P.devise)}</dd>
              <dt>Jours d’exploitation / semaine</dt><dd>${fmtN(P.joursSem)}</dd>
              <dt>Date de référence</dt><dd>${P.refDate ? fmtD(P.refDate) : `Aujourd’hui (${fmtD(ctx.ref)})`}</dd>
            </dl>
            <h3 style="font-size:13px;margin:18px 0 10px;color:var(--text-2)">Ce qui a été reconnu</h3>
            <ul class="check-list">${checks}</ul>
          </div></section>
      </div>
      <section class="card" aria-labelledby="obj-t"><div class="card-h"><div><h2 id="obj-t">Objectifs et seuils</h2><p>Définis dans la feuille Paramètres du fichier Excel (vert = objectif atteint, orange = entre objectif et seuil, rouge = au-delà du seuil)</p></div></div>
        <div class="card-b flush">${dataTable({ id: 'objTable', caption: 'Objectifs et seuils', cols: tcols, rows: ctx.res.kpis })}</div></section>`;
    return {
      html,
      mount() {
        $('file2').onchange = (e) => e.target.files[0] && loadFile(e.target.files[0]);
        $('dlTemplate').onclick = () => XLSX.writeFile(C.buildWorkbook(XLSX, null), 'modele_gmao_leon.xlsx');
        $('dlExport').onclick = () => XLSX.writeFile(C.buildWorkbook(XLSX, S.data), S.isDemo ? 'demo_gmao_leon.xlsx' : 'export_gmao_leon.xlsx');
        $('dlDemo').onclick = () => { loadDemo(); toast('Données de démonstration rechargées.'); };
        wireTable('objTable');
      },
    };
  }

  // ---------- equipment detail sheet ----------
  function sheetHtml(ctx, code) {
    const a = ctx.assets.find((x) => x.code === code) || null;
    const e = ctx.eqBy[code];
    if (!e) return `<div class="sheet-head"><div><h2 id="sheetTitle">Équipement introuvable</h2><p>${esc(code)}</p></div><button class="icon-btn" type="button" data-close autofocus aria-label="Fermer">${icon('x')}</button></div>`;
    const st = a ? a.st : assetStatus(e, ctx.ref);
    const ref = ctx.ref;
    const pannes = S.data.pannes.filter((p) => p.code === code).sort((x, y) => y.date - x.date);
    const inter = S.data.interventions.filter((i) => i.code === code && (i.parts || i.partRef)).sort((x, y) => y.date - x.date);
    const upcoming = S.data.preventif.filter((t) => t.code === code && t.done == null && t.due != null).sort((x, y) => x.due - y.due).slice(0, 6);
    const ots = S.data.ot.filter((o) => o.code === code && !isClosedOT(o));
    const s = a && a.service;
    const dueAnnual = s ? (s.annualState === 'late' ? badge('bad', 'En retard') : s.annualState === 'soon' ? badge('warn', 'Bientôt') : s.annualState === 'ok' ? badge('good', 'À jour') : badge('neutral', 'Date inconnue')) : '';
    const since = e.commissioned != null ? toDate(e.commissioned).getUTCFullYear() : null;
    return `
      <div class="sheet-head">
        <span class="asset-ico">${icon('monitor-smartphone')}</span>
        <div><h2 id="sheetTitle">${esc(code)} · ${esc(e.model || e.name)}</h2>
          <p>${esc([e.zone, e.line].filter(Boolean).join(' · '))}${e.serial ? ` · N° série ${esc(e.serial)}` : ''}</p>
          <div class="badges">${badge(st.key, st.label)}${e.crit ? `<span class="badge neutral">${esc('Criticité : ' + e.crit)}</span>` : ''}${since ? `<span class="badge neutral">En service depuis ${since}</span>` : ''}${e.hours != null ? `<span class="badge neutral">${fmtN(e.hours)} h de fonctionnement</span>` : ''}</div></div>
        <button class="icon-btn" type="button" data-close autofocus aria-label="Fermer la fiche">${icon('x')}</button>
      </div>
      <div class="sheet-body">
        ${ots.length ? `<div class="notice warn">${icon('wrench')}<div><strong>${plural(ots.length, 'ordre de travail ouvert', 'ordres de travail ouverts')}</strong><br>${ots.map((o) => `${esc(o.id)} · ${esc(o.desc)} · ${esc(o.status)}${o.planned != null ? ` (prévu le ${fmtD(o.planned)})` : ''}`).join('<br>')}</div></div>` : ''}
        <section class="card"><div class="sheet-kpis mini-stats">
          ${miniPlain('activity', 'Disponibilité', a ? fmtPct(a.dispo, 2).replace(' %', '') : '—', '%', 'Sur la période')}
          ${miniPlain('triangle-alert', 'Pannes', a ? fmtN(a.pannes) : '—', '', 'Sur la période')}
          ${miniPlain('timer', 'MTBF', a ? fmtN(a.mtbf) : '—', 'h', 'Entre deux pannes')}
          ${miniPlain('wallet', 'Coût', a ? fmtN(a.cost) : '—', S.data.params.devise, 'Sur la période')}
        </div></section>
        ${s ? `<div class="sheet-section"><h3>Échéances constructeur</h3><div class="due-cards">
          <div class="card due"><span class="k">Maintenance annuelle + CST</span><span class="v">${fmtD(s.nextAnnual)}</span>${dueAnnual}<span class="h">${s.lastAnnual != null ? `Dernière réalisée le ${fmtD(s.lastAnnual)}` : 'Aucune maintenance annuelle enregistrée'}</span></div>
          <div class="card due"><span class="k">Révision générale</span><span class="v">${s.overhaul ? fmtD(s.overhaul.date) : '—'}</span>${s.overhaul ? `<span class="h">Révision ${esc(s.overhaul.kind)}${s.overhaul.hoursNext != null ? `, ou à ${fmtN(s.overhaul.hoursNext)} h (dans ${fmtN(s.overhaul.hoursLeft)} h) si atteint avant` : ''}</span>` : '<span class="h">Date de mise en service inconnue</span>'}</div>
        </div></div>` : ''}
        <div class="sheet-section"><h3>Prochaines tâches préventives</h3><section class="card">${upcoming.length ? `<ul class="feed">${upcoming.map((t) => `<li><span class="d">${fmtD(t.due)}</span><div><strong>${esc(t.op)}</strong><span class="sub">${esc(t.freq || '')}${t.resp ? ` · ${esc(t.resp)}` : ''}</span></div><span class="r">${t.due < ref ? badge('bad', 'En retard') : badge('info', `J-${t.due - ref}`)}</span></li>`).join('')}</ul>` : `<div class="empty">Aucune tâche planifiée.</div>`}</section></div>
        <div class="sheet-section"><h3>Historique des pannes</h3><section class="card">${pannes.length ? `<ul class="feed">${pannes.slice(0, 8).map((p) => `<li><span class="d">${fmtD(p.date)}</span><div><strong>${esc(p.desc)}</strong><span class="sub">${esc(p.cause)}${p.alarm ? ` · alarme ${esc(p.alarm)}` : ''}${p.resp ? ` · ${esc(p.resp)}` : ''}</span></div><span class="r">${fmtN(p.downtime, 1)} h</span></li>`).join('')}</ul>${pannes.length > 8 ? `<div class="table-foot">${plural(pannes.length - 8, 'autre panne', 'autres pannes')} dans le registre</div>` : ''}` : `<div class="empty">${icon('circle-check')}Aucune panne enregistrée.</div>`}</section></div>
        <div class="sheet-section"><h3>Pièces remplacées</h3><section class="card">${inter.length ? `<ul class="feed">${inter.slice(0, 8).map((i) => `<li><span class="d">${fmtD(i.date)}</span><div><strong>${esc(i.parts || i.partRef)}</strong><span class="sub">${i.partRef ? `Réf. ${esc(i.partRef)} · ` : ''}${esc(C.typeGroup(i.type))}</span></div><span class="r">${fmtMoney(i.costParts)}</span></li>`).join('')}</ul>` : `<div class="empty">Aucune pièce enregistrée.</div>`}</section></div>
      </div>`;
  }

  function openSheet(code) {
    const dlg = $('sheet');
    $('sheetBody').innerHTML = sheetHtml(S.ctx, code);
    paintIcons();
    $('sheetBody').querySelector('[data-close]').onclick = () => dlg.close();
    if (!dlg.open) dlg.showModal();
    $('sheetBody').querySelector('[data-close]').focus();
  }

  // ---------- routing & render ----------
  function parseRoute() {
    const parts = location.hash.replace(/^#\/?/, '').split('/');
    const page = PAGES[parts[0]] ? parts[0] : 'apercu';
    return { page, param: parts[1] ? decodeURIComponent(parts[1]) : null };
  }

  function update() {
    S.ctx = null;
    buildCtx();
    renderChrome();
    renderPage(true);
  }

  function renderChrome() {
    const ctx = S.ctx;
    const res = ctx.res;
    $('periodText').textContent = `Du ${fmtD(res.from)} au ${fmtD(res.toEff)} · ${plural(res.days, 'jour', 'jours')}`;
    const set = (k, v, warn) => document.querySelectorAll(`[data-count="${k}"]`).forEach((el) => { el.textContent = v ? String(v) : ''; el.classList.toggle('warn', !!warn); });
    set('actions', ctx.actions.filter((a) => a.sev === 'bad').length || ctx.actions.filter((a) => a.sev === 'warn').length, !ctx.actions.some((a) => a.sev === 'bad'));
    set('down', ctx.assets.filter((a) => a.st.key === 'bad').length);
    set('prevLate', res.prevLate.length, true);
    set('otLate', ctx.kpi.otRetard.value, true);
    $('notices').querySelectorAll('.notice.demo').forEach((n) => n.remove());
    if (S.isDemo && !S.ui.demoDismissed) {
      const el = document.createElement('div');
      el.className = 'notice demo';
      el.innerHTML = `${icon('info')}<div><strong>Mode démonstration</strong> : 8 stations LM leon fictives. Importez votre fichier Excel (bouton « Importer Excel ») pour travailler sur vos données réelles.</div><button class="icon-btn x" type="button" aria-label="Fermer le message">${icon('x')}</button>`;
      el.querySelector('.x').onclick = () => { S.ui.demoDismissed = true; el.remove(); };
      $('notices').prepend(el);
    }
    if (!S.isDemo && S.report && S.report.warnings.length && !S.ui.warnShown) {
      S.ui.warnShown = true;
      S.report.warnings.forEach((w) => notice('warn', w));
    }
  }

  function renderPage(keepFocus) {
    const route = parseRoute();
    const pageChanged = route.page !== S.route.page;
    S.route = route;
    const page = PAGES[route.page];
    S.charts.forEach((c) => c.destroy());
    S.charts = [];
    for (const k of Object.keys(tables)) delete tables[k];
    const scrollY = window.scrollY;
    const { html, mount } = page.render(S.ctx);
    $('content').innerHTML = html;
    document.title = `${page.title} · GMAO leon`;
    $('pageTitle').textContent = page.title;
    $('pageDesc').textContent = page.desc;
    $('crumbs').innerHTML = `<a href="#/apercu">${icon('house')}<span class="sr-only">Accueil</span></a>${icon('chevron-right')}<span>${esc(page.title)}</span>`;
    document.querySelectorAll('#sideNav a').forEach((a) => (a.dataset.page === route.page ? a.setAttribute('aria-current', 'page') : a.removeAttribute('aria-current')));
    paintIcons();
    mount();
    if (pageChanged && !keepFocus) {
      window.scrollTo(0, 0);
      $('content').focus({ preventScroll: true });
    } else if (keepFocus === true) window.scrollTo(0, scrollY);
    if (route.page === 'equipements' && route.param) openSheet(route.param);
    else if ($('sheet').open) $('sheet').close();
  }

  // ---------- shell behaviour ----------
  const mqMobile = matchMedia('(max-width: 1023px)');
  function openNav() {
    const side = $('sidebar');
    side.classList.add('open');
    side.setAttribute('role', 'dialog');
    side.setAttribute('aria-modal', 'true');
    $('main').inert = true;
    $('scrim').hidden = false;
    $('menuBtn').setAttribute('aria-expanded', 'true');
    side.querySelector('#sideClose').focus();
  }
  function closeNav(returnFocus = true) {
    const side = $('sidebar');
    if (!side.classList.contains('open')) return;
    side.classList.remove('open');
    side.removeAttribute('role');
    side.removeAttribute('aria-modal');
    $('main').inert = false;
    $('scrim').hidden = true;
    $('menuBtn').setAttribute('aria-expanded', 'false');
    if (returnFocus) $('menuBtn').focus();
  }

  function setCollapsed(on) {
    $('app').classList.toggle('collapsed', on);
    $('btnCollapse').setAttribute('aria-pressed', String(on));
    $('btnCollapse').querySelector('span').textContent = on ? 'Déplier' : 'Réduire';
    document.querySelectorAll('#sideNav a').forEach((a) => (on ? a.setAttribute('title', a.textContent.trim()) : a.removeAttribute('title')));
    try { localStorage.setItem('gm-collapsed', on ? '1' : '0'); } catch (e) { /* ignore */ }
  }

  function init() {
    const img = $('logoImg');
    const miss = () => $('brandLogo').classList.add('missing');
    img.addEventListener('error', miss);
    if (img.complete && !img.naturalWidth) miss();

    try {
      const t = localStorage.getItem('gm-theme');
      if (t) document.documentElement.dataset.theme = t;
      if (localStorage.getItem('gm-collapsed') === '1') setCollapsed(true);
    } catch (e) { /* storage unavailable */ }

    $('btnTheme').onclick = () => {
      const cur = document.documentElement.dataset.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
      document.documentElement.dataset.theme = cur === 'dark' ? 'light' : 'dark';
      try { localStorage.setItem('gm-theme', document.documentElement.dataset.theme); } catch (e) { /* ignore */ }
      renderPage(true);
    };
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => renderPage(true));
    $('btnCollapse').onclick = () => setCollapsed(!$('app').classList.contains('collapsed'));
    $('menuBtn').onclick = openNav;
    $('sideClose').onclick = () => closeNav();
    $('scrim').onclick = () => closeNav();
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && $('sidebar').classList.contains('open')) closeNav(); });
    $('sideNav').addEventListener('click', (e) => { if (e.target.closest('a') && mqMobile.matches) closeNav(false); });
    mqMobile.addEventListener('change', () => closeNav(false));

    $('fPeriod').onchange = (e) => setPeriod(e.target.value);
    $('fFrom').onchange = $('fTo').onchange = () => {
      const f = C.toSerial($('fFrom').value), t = C.toSerial($('fTo').value);
      if (f == null || t == null || f > t) return;
      S.from = f; S.to = t;
      update();
    };
    $('fZone').onchange = (e) => { S.zone = e.target.value; update(); };
    $('fModel').onchange = (e) => { S.model = e.target.value; update(); };
    $('file').onchange = (e) => e.target.files[0] && loadFile(e.target.files[0]);
    $('btnPrint').onclick = () => window.print();

    const sheet = $('sheet');
    sheet.addEventListener('close', () => {
      const r = parseRoute();
      if (r.page === 'equipements' && r.param) history.replaceState(null, '', '#/equipements');
    });
    if (!('closedBy' in HTMLDialogElement.prototype)) {
      let downOutside = false;
      const outside = (ev) => { const r = sheet.getBoundingClientRect(); return ev.clientX < r.left || ev.clientX > r.right || ev.clientY < r.top || ev.clientY > r.bottom; };
      sheet.addEventListener('pointerdown', (ev) => (downOutside = ev.target === sheet && outside(ev)));
      sheet.addEventListener('pointerup', (ev) => { if (downOutside && ev.target === sheet && outside(ev)) sheet.close(); downOutside = false; });
    }

    window.addEventListener('hashchange', () => {
      const r = parseRoute();
      if (r.page === S.route.page && r.page === 'equipements') {
        S.route = r;
        if (r.param) openSheet(r.param);
        else if (sheet.open) sheet.close();
        return;
      }
      renderPage(false);
    });

    let depth = 0;
    window.addEventListener('dragenter', (e) => { if (!e.dataTransfer.types.includes('Files')) return; e.preventDefault(); depth++; $('drop').classList.add('on'); });
    window.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; $('drop').classList.remove('on'); } });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault(); depth = 0; $('drop').classList.remove('on');
      const f = e.dataTransfer.files[0];
      if (f) loadFile(f);
    });

    if (!location.hash) history.replaceState(null, '', '#/apercu');
    S.route = parseRoute();
    loadDemo();
  }

  init();
  window.GMApp = { loadBuffer };
})();
