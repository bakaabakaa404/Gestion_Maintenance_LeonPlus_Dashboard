/* Dashboard UI: file import, filters, KPI tiles, charts (Chart.js) and tables. Logic lives in core.js. */
(function () {
  'use strict';
  const C = window.GMCore;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const state = { data: null, report: null, source: '', isDemo: false, f: {}, asTable: new Set(), charts: [] };

  // ---------- formatting ----------
  const nf = (d) => new Intl.NumberFormat('fr-FR', { minimumFractionDigits: d, maximumFractionDigits: d });
  const fmtN = (v, d = 0) => (v == null || !isFinite(v) ? '—' : nf(d).format(v));
  const fmtPct = (v, d = 1) => (v == null || !isFinite(v) ? '—' : `${nf(d).format(v * 100)} %`);
  const fmtD = (s) => (s == null ? '—' : C.fmtDate(s));
  const isoFromSerial = (s) => new Date((s - 25569) * 86400000).toISOString().slice(0, 10);
  const serialFromIso = (v) => (v ? C.toSerial(v) : null);
  const css = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

  const KPI_FMT = {
    pannes: (v) => [fmtN(v), ''],
    dispo: (v) => [v == null ? '—' : nf(2).format(v * 100), '%'],
    prev: (v) => [v == null ? '—' : nf(1).format(v * 100), '%'],
    mtbf: (v) => [fmtN(v, 0), 'h'],
    mttr: (v) => [fmtN(v, 2), 'h'],
    arret: (v) => [fmtN(v, 1), 'h'],
    cout: (v, dev) => [fmtN(v, 0), dev],
    otRetard: (v) => [fmtN(v), 'OT'],
  };
  const targetText = (k, t, dev) => {
    const sign = t.higher ? '≥' : '≤';
    const val = k.unit === '%' ? fmtPct(t.obj, t.obj > 0.99 ? 1 : 0) : k.key === 'cout' ? `${fmtN(t.obj)} ${dev}` : `${fmtN(t.obj, k.key === 'mttr' ? 1 : 0)}${k.unit === 'h' ? ' h' : ''}`;
    return `Objectif : ${sign} ${val}`;
  };
  const STATUS = {
    good: ['✓', 'Objectif atteint'],
    warn: ['!', 'À surveiller'],
    bad: ['✕', 'Hors objectif'],
    na: ['–', 'Non calculable'],
  };
  const pill = (st, text) => `<span class="pill ${st}"><span class="dot" aria-hidden="true">${STATUS[st][0]}</span>${esc(text || STATUS[st][1])}</span>`;

  // ---------- data loading ----------
  function setData(data, report, source, isDemo) {
    state.data = data;
    state.report = report;
    state.source = source;
    state.isDemo = isDemo;
    state.f = { eq: '', model: '', zone: '', crit: '', type: 'Tous' };
    fillFilters();
    const range = C.dataRange(data);
    const ref = refDate();
    const spanMonths = range.min != null ? (ref - range.min) / 30.4 : 0;
    applyQuick(spanMonths > 13 ? '12m' : 'all');
  }

  const refDate = () => (state.data.params.refDate != null ? state.data.params.refDate : C.todaySerial());

  function loadDemo() {
    setData(C.demoData(), null, 'Données de démonstration (fictives)', true);
  }

  function loadFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => loadBuffer(e.target.result, file.name);
    reader.readAsArrayBuffer(file);
  }

  function loadBuffer(buf, name) {
    const file = { name };
    {
      try {
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
        const { data, report } = C.parseWorkbook(XLSX, wb);
        if (!data.pannes.length && !data.interventions.length && !data.equipements.length) {
          banner('warn', `<strong>${esc(file.name)}</strong> : aucune feuille reconnue (registre des pannes, interventions, équipements…). Téléchargez le modèle Excel pour voir les en-têtes attendus.`);
          return;
        }
        setData(data, report, file.name, false);
      } catch (err) {
        banner('warn', `Lecture impossible de <strong>${esc(file.name)}</strong> : ${esc(err.message)}`);
      }
    }
  }

  function banner(kind, html) {
    const el = document.createElement('div');
    el.className = `banner ${kind}`;
    el.innerHTML = `<div>${html}</div><button class="x" type="button" aria-label="Fermer">✕</button>`;
    el.querySelector('.x').onclick = () => el.remove();
    $('banners').appendChild(el);
  }

  // ---------- filters ----------
  function fillFilters() {
    const o = C.filterOptions(state.data);
    const opts = (sel, items, all) => {
      sel.innerHTML = `<option value="">${all}</option>` + items.map((i) => (typeof i === 'string' ? { value: i, label: i } : i)).map((i) => `<option value="${esc(i.value)}">${esc(i.label)}</option>`).join('');
    };
    opts($('fEq'), o.equipements, 'Tous les équipements');
    opts($('fModel'), o.models, 'Tous');
    opts($('fZone'), o.zones, 'Tous');
    opts($('fCrit'), o.crits, 'Toutes');
    $('fType').innerHTML = ['Tous'].concat(o.types).map((t) => `<option>${esc(t)}</option>`).join('');
  }

  function applyQuick(q) {
    const ref = refDate();
    const range = C.dataRange(state.data);
    let from, to = ref;
    if (q === '12m') from = C.addMonths(ref, -12) + 1;
    else if (q === 'ytd') from = C.toSerial(`${new Date((ref - 25569) * 86400000).getUTCFullYear()}-01-01`);
    else {
      from = range.min != null ? range.min : ref - 365;
      if (range.max != null && range.max > to) to = range.max;
    }
    $('fFrom').value = isoFromSerial(from);
    $('fTo').value = isoFromSerial(to);
    document.querySelectorAll('#quick button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.q === q)));
    render();
  }

  function readFilters() {
    state.f.from = serialFromIso($('fFrom').value);
    state.f.to = serialFromIso($('fTo').value);
    state.f.eq = $('fEq').value;
    state.f.model = $('fModel').value;
    state.f.zone = $('fZone').value;
    state.f.crit = $('fCrit').value;
    state.f.type = $('fType').value || 'Tous';
  }

  // ---------- render ----------
  function render() {
    readFilters();
    if (state.f.from == null || state.f.to == null || state.f.from > state.f.to) return;
    const res = C.compute(state.data, state.f);
    state.res = res;
    renderHeader(res);
    renderKpis(res);
    renderCharts(res);
    renderTables(res);
    renderDiag();
  }

  function renderHeader(res) {
    const n = state.data;
    $('sourceLine').textContent = `${state.source} · ${n.equipements.length} équipement(s) · ${n.pannes.length} panne(s) · ${n.interventions.length} intervention(s)`;
    $('periodNote').textContent = `Période analysée : ${res.days} jour(s), du ${fmtD(res.from)} au ${fmtD(res.toEff)} · date de référence ${fmtD(res.ref)} · ${state.data.params.joursSem} jours d'exploitation / semaine`;
    $('banners').querySelectorAll('.banner.demo, .banner.import').forEach((b) => b.remove());
    if (state.isDemo) {
      const el = document.createElement('div');
      el.className = 'banner demo';
      el.innerHTML = `<div><strong>Données fictives de démonstration</strong> : 8 stations LM leon (leon plus, leon, leon mri). Les codes alarme et intervalles suivent le manuel leon plus ; les références « DEMO-REF » sont à remplacer par celles de la liste officielle des pièces de rechange. Importez votre fichier Excel (modèle SAMQ ou modèle vierge ci-dessus) pour analyser vos données.</div>`;
      $('banners').prepend(el);
    } else if (state.report && state.report.warnings.length) {
      const el = document.createElement('div');
      el.className = 'banner warn import';
      el.innerHTML = `<div>${state.report.warnings.map(esc).join('<br>')}</div>`;
      $('banners').prepend(el);
    }
    $('leonRef').style.display = res.hasLeon ? '' : 'none';
  }

  function renderKpis(res) {
    $('kpis').innerHTML = res.kpis
      .map((k) => {
        const [v, unit] = KPI_FMT[k.key](k.value, res.devise);
        return `<div class="card kpi">
          <div class="name">${esc(k.label)}</div>
          <div class="value">${esc(v)}${unit ? `<span class="unit">${esc(unit)}</span>` : ''}</div>
          <div class="detail">${esc(k.detail)}</div>
          <div class="foot">${pill(k.status)}<span class="target">${esc(targetText(k, k.target, res.devise))}</span></div>
        </div>`;
      })
      .join('');
  }

  // ---------- charts ----------
  const trunc = (s, n = 30) => (s.length > n ? `${s.slice(0, n - 1)}…` : s);

  function chartSpecs(res) {
    const top = (arr, n = 10) => arr.slice(0, n);
    const dev = res.devise;
    const causes = top(res.causes);
    const n80 = res.causes.findIndex((c) => c.cumPct >= 0.8) + 1;
    const dispoTarget = res.kpis.find((k) => k.key === 'dispo').target.obj;
    const costs = top(res.costByEq);
    const resp = top(res.byResp);
    const respTypes = res.respTypes.filter((t) => resp.some((r) => r.hours[t] > 0));
    const S = ['--series-1', '--series-2', '--series-3', '--series-4'];

    return [
      {
        id: 'monthly', title: 'Nombre de pannes par mois', sub: 'Pannes incluses dans les KPI',
        kind: 'bar', labels: res.monthly.map((m) => m.label),
        series: [{ label: 'Pannes', data: res.monthly.map((m) => m.pannes), color: S[0] }],
        fmt: (v) => fmtN(v), extra: (i) => `Arrêt : ${fmtN(res.monthly[i].downtime, 1)} h`,
      },
      {
        id: 'dispo', title: 'Disponibilité technique par mois', sub: `Ligne de référence : objectif ${fmtPct(dispoTarget, 1)}`,
        kind: 'line', labels: res.monthly.map((m) => m.label),
        series: [{ label: 'Disponibilité', data: res.monthly.map((m) => m.dispo), color: S[0] }],
        target: { label: 'Objectif', value: dispoTarget }, pct: true, fmt: (v) => fmtPct(v, 2),
      },
      {
        id: 'causes', title: 'Principales causes des pannes correctives',
        sub: res.causes.length ? `Top 10 · ${n80} cause(s) expliquent 80 % des pannes` : '',
        kind: 'hbar', labels: causes.map((c) => c.label), tall: true,
        series: [{ label: 'Pannes', data: causes.map((c) => c.value), color: S[0] }],
        fmt: (v) => fmtN(v), extra: (i) => `Cumul : ${fmtPct(causes[i].cumPct, 0)} · arrêt ${fmtN(causes[i].downtime, 1)} h`,
        tableExtra: { head: ['Cumul', 'Arrêt (h)'], row: (i) => [fmtPct(causes[i].cumPct, 0), fmtN(causes[i].downtime, 1)] },
      },
      {
        id: 'families', title: 'Pannes par sous-ensemble', sub: 'Famille de panne / module concerné',
        kind: 'hbar', labels: top(res.families).map((c) => c.label), tall: true,
        series: [{ label: 'Pannes', data: top(res.families).map((c) => c.value), color: S[0] }],
        fmt: (v) => fmtN(v),
      },
      {
        id: 'downEq', title: "Temps d'arrêt sur pannes par équipement", sub: 'Top 10, en heures',
        kind: 'hbar', labels: top(res.downByEq).map((c) => c.label),
        series: [{ label: 'Arrêt (h)', data: top(res.downByEq).map((c) => c.value), color: S[0] }],
        fmt: (v) => `${fmtN(v, 1)} h`,
      },
      {
        id: 'crit', title: 'Pannes par niveau de criticité', sub: "Criticité définie dans la liste des équipements",
        kind: 'bar', labels: res.critData.map((c) => c.label),
        series: [{ label: 'Pannes', data: res.critData.map((c) => c.pannes), color: S[0] }],
        fmt: (v) => fmtN(v), extra: (i) => `Arrêt : ${fmtN(res.critData[i].downtime, 1)} h`,
        tableExtra: { head: ['Arrêt (h)'], row: (i) => [fmtN(res.critData[i].downtime, 1)] },
      },
      {
        id: 'cost', title: 'Coût de maintenance par équipement', sub: `Top 10, en ${dev} · hors interventions annulées`,
        kind: 'hbar', stacked: true, labels: costs.map((c) => c.label),
        series: [
          { label: "Main-d'œuvre", data: costs.map((c) => c.mo), color: S[0] },
          { label: 'Pièces', data: costs.map((c) => c.parts), color: S[1] },
          { label: 'Autres', data: costs.map((c) => c.other), color: S[2] },
        ].filter((s) => s.data.some((v) => v > 0)),
        fmt: (v) => `${fmtN(v)} ${dev}`, footer: (i) => `Total : ${fmtN(costs[i].total)} ${dev}`,
      },
      {
        id: 'resp', title: 'Charge de travail par responsable', sub: "Heures d'intervention par type de maintenance",
        kind: 'hbar', stacked: true, labels: resp.map((r) => r.label),
        series: respTypes.map((t) => ({ label: t, data: resp.map((r) => r.hours[t]), color: S[res.respTypes.indexOf(t)] })),
        fmt: (v) => `${fmtN(v, 1)} h`, footer: (i) => `Total : ${fmtN(resp[i].total, 1)} h · ${resp[i].count} intervention(s)`,
      },
    ];
  }

  function renderCharts(res) {
    state.charts.forEach((c) => c.destroy());
    state.charts = [];
    const host = $('charts');
    host.innerHTML = '';
    const grid = css('--grid'), muted = css('--text-muted'), text = css('--text-primary'), text2 = css('--text-secondary');
    const surface = css('--surface-1'), border = css('--border');
    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;

    for (const spec of chartSpecs(res)) {
      const card = document.createElement('div');
      card.className = 'card chart-card' + (state.asTable.has(spec.id) ? ' as-table' : '');
      const hasData = spec.labels.length && spec.series.some((s) => s.data.some((v) => v != null && v !== 0));
      const legendItems = spec.series.length > 1 || spec.target
        ? spec.series.map((s) => `<span><i style="background:${css(s.color)}"></i>${esc(s.label)}</span>`).join('') +
          (spec.target ? `<span><i class="line" style="background:${css('--ref-line')}"></i>${esc(spec.target.label)}</span>` : '')
        : '';
      card.innerHTML = `<header><div><h3>${esc(spec.title)}</h3>${spec.sub ? `<div class="sub">${esc(spec.sub)}</div>` : ''}</div>
          <button class="toggle" type="button" aria-pressed="${state.asTable.has(spec.id)}">${state.asTable.has(spec.id) ? 'Graphique' : 'Tableau'}</button></header>
        ${legendItems ? `<div class="legend">${legendItems}</div>` : ''}
        <div class="canvas-box${spec.tall ? ' tall' : ''}">${hasData ? '<canvas role="img"></canvas>' : '<div class="empty">Aucune donnée sur la période et les filtres choisis.</div>'}</div>
        <div class="chart-table"></div>`;
      host.appendChild(card);
      card.querySelector('.chart-table').innerHTML = chartTable(spec);
      card.querySelector('.toggle').onclick = (e) => {
        const on = !state.asTable.has(spec.id);
        on ? state.asTable.add(spec.id) : state.asTable.delete(spec.id);
        card.classList.toggle('as-table', on);
        e.currentTarget.textContent = on ? 'Graphique' : 'Tableau';
        e.currentTarget.setAttribute('aria-pressed', String(on));
      };
      if (!hasData) continue;

      const canvas = card.querySelector('canvas');
      canvas.setAttribute('aria-label', `${spec.title}. Utilisez le bouton Tableau pour lire les valeurs.`);
      const horizontal = spec.kind === 'hbar';
      const isLine = spec.kind === 'line';
      const datasets = spec.series.map((s, si) => {
        const color = css(s.color);
        if (isLine)
          return { label: s.label, data: s.data, borderColor: color, backgroundColor: color, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, pointBackgroundColor: color, pointBorderColor: surface, pointBorderWidth: 2, tension: 0, spanGaps: true };
        const last = si === spec.series.length - 1;
        return {
          label: s.label, data: s.data, backgroundColor: color, hoverBackgroundColor: color,
          borderColor: surface, borderWidth: spec.stacked ? (horizontal ? { right: 2 } : { top: 2 }) : 0,
          borderRadius: spec.stacked ? (last ? { topRight: 4, bottomRight: 4, topLeft: 0, bottomLeft: 0 } : 0) : 4,
          borderSkipped: spec.stacked ? false : 'start',
          maxBarThickness: horizontal ? 18 : 28, categoryPercentage: 0.8, barPercentage: 0.9,
        };
      });
      if (spec.target) {
        datasets.push({ label: spec.target.label, data: spec.labels.map(() => spec.target.value), borderColor: css('--ref-line'), borderWidth: 1, borderDash: [4, 3], pointRadius: 0, pointHoverRadius: 0, fill: false });
      }
      const valueAxis = {
        beginAtZero: !spec.pct, stacked: !!spec.stacked,
        grid: { color: grid, drawTicks: false }, border: { display: false },
        ticks: { color: muted, font: { size: 11 }, padding: 6, maxTicksLimit: 6, callback: (v) => (spec.pct ? fmtPct(v, 1) : fmtN(v)) },
      };
      if (spec.pct) {
        const vals = spec.series[0].data.filter((v) => v != null).concat([spec.target.value]);
        valueAxis.min = Math.max(0, Math.floor((Math.min(...vals) - 0.003) * 200) / 200);
        valueAxis.max = 1;
      }
      const catAxis = {
        stacked: !!spec.stacked, grid: { display: false }, border: { color: grid },
        ticks: { color: text2, font: { size: 11 }, autoSkip: !horizontal, maxRotation: 0, callback: function (v) { return trunc(this.getLabelForValue(v), horizontal ? 32 : 14); } },
      };
      const chart = new Chart(canvas, {
        type: isLine ? 'line' : 'bar',
        data: { labels: spec.labels, datasets },
        options: {
          responsive: true, maintainAspectRatio: false, animation: { duration: 200 },
          indexAxis: horizontal ? 'y' : 'x',
          interaction: { mode: 'index', intersect: false, axis: horizontal ? 'y' : 'x' },
          layout: { padding: { right: 8 } },
          scales: horizontal ? { x: valueAxis, y: catAxis } : { x: catAxis, y: valueAxis },
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: surface, titleColor: text, bodyColor: text2, footerColor: text2, borderColor: border, borderWidth: 1,
              padding: 10, boxPadding: 4, usePointStyle: true, titleFont: { weight: '600' }, footerFont: { weight: '500' },
              filter: (item) => !(spec.target && item.datasetIndex === spec.series.length),
              callbacks: {
                title: (items) => items[0].label,
                label: (item) => ` ${item.dataset.label} : ${spec.fmt(item.raw)}`,
                afterBody: (items) => (spec.extra ? spec.extra(items[0].dataIndex) : ''),
                footer: (items) => (spec.footer ? spec.footer(items[0].dataIndex) : spec.target ? `Objectif : ${spec.fmt(spec.target.value)}` : ''),
              },
            },
          },
        },
      });
      state.charts.push(chart);
    }
  }

  function chartTable(spec) {
    const head = [spec.kind === 'line' || spec.id === 'monthly' ? 'Mois' : 'Libellé']
      .concat(spec.series.map((s) => s.label))
      .concat(spec.tableExtra ? spec.tableExtra.head : []);
    const rows = spec.labels.map((l, i) => [esc(l)].concat(spec.series.map((s) => esc(spec.fmt(s.data[i])))).concat(spec.tableExtra ? spec.tableExtra.row(i).map(esc) : []));
    if (!rows.length) return '<div class="empty">Aucune donnée.</div>';
    return `<table><thead><tr>${head.map((h, i) => `<th class="${i ? 'num' : ''}">${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows
      .map((r) => `<tr>${r.map((c, i) => `<td class="${i ? 'num' : ''}">${c}</td>`).join('')}</tr>`)
      .join('')}</tbody></table>`;
  }

  // ---------- tables ----------
  function table(cols, rows, empty = 'Aucun élément.') {
    if (!rows.length) return `<div class="empty">${esc(empty)}</div>`;
    return `<table><thead><tr>${cols.map((c) => `<th class="${c.num ? 'num' : ''}">${esc(c.label)}</th>`).join('')}</tr></thead><tbody>${rows
      .map((r) => `<tr>${cols.map((c, i) => `<td class="${c.num ? 'num' : c.nowrap ? 'nowrap' : ''}">${r[i] == null || r[i] === '' ? '<span class="muted">—</span>' : r[i]}</td>`).join('')}</tr>`)
      .join('')}</tbody></table>`;
  }

  const critTag = (c) => (c ? `<span class="crit-${esc(c)}">${esc(c)}</span>` : '');
  const eqName = (code) => {
    const e = state.data.equipements.find((x) => x.code === code);
    return e ? `<strong>${esc(code)}</strong> <span class="muted">${esc(e.model || e.name)}</span>` : esc(code);
  };

  function renderTables(res) {
    const dev = res.devise;
    const anyLeon = res.fleet.some((e) => e.service);
    const annualPill = (s) => {
      if (!s) return '';
      const map = { late: ['bad', 'En retard'], soon: ['warn', '≤ 30 jours'], ok: ['good', 'À jour'], unknown: ['na', 'Non renseignée'] };
      const [st, label] = map[s.annualState];
      return `${s.nextAnnual != null ? `${fmtD(s.nextAnnual)}<br>` : ''}${pill(st, label)}`;
    };
    const overhaul = (s) => {
      if (!s || !s.overhaul) return '';
      const o = s.overhaul;
      return `${fmtD(o.date)} <span class="muted">(${o.kind})</span>${o.hoursNext != null ? `<br><span class="muted">ou à ${fmtN(o.hoursNext)} h (${o.hoursKind}) · dans ${fmtN(o.hoursLeft)} h</span>` : ''}`;
    };
    const fcols = [
      { label: 'Équipement' }, { label: 'N° de série' }, { label: 'Service / salle' }, { label: 'Criticité' },
      { label: 'Compteur (h)', num: true }, { label: 'Pannes', num: true }, { label: 'Arrêt (h)', num: true },
      { label: 'Disponibilité', num: true }, { label: 'MTBF (h)', num: true }, { label: `Coût (${dev})`, num: true }, { label: 'Dernière panne', nowrap: true },
    ];
    if (anyLeon) fcols.push({ label: 'Prochaine maintenance annuelle + CST', nowrap: true }, { label: 'Prochaine révision générale' });
    $('fleetTable').innerHTML = table(
      fcols,
      res.fleet.map((e) => {
        const row = [eqName(e.code), esc(e.serial), `${esc(e.zone)}${e.line ? ` <span class="muted">· ${esc(e.line)}</span>` : ''}`, critTag(e.crit),
          fmtN(e.hours), fmtN(e.pannes), fmtN(e.downtime, 1), fmtPct(e.dispo, 2), fmtN(e.mtbf), fmtN(e.cost), fmtD(e.lastPanne)];
        if (anyLeon) row.push(annualPill(e.service), overhaul(e.service));
        return row;
      }),
      'Aucun équipement dans le périmètre filtré.'
    );
    $('fleetSub').textContent = anyLeon
      ? 'Échéances calculées selon le manuel : annuelle (12 mois), révision 3 ans / 10 000 h et 6 ans / 20 000 h'
      : `${res.fleet.length} équipement(s)`;

    $('partsTable').innerHTML = table(
      [{ label: 'Référence' }, { label: 'Composant / pièce' }, { label: 'Quantité', num: true }, { label: `Coût pièces (${dev})`, num: true }, { label: 'Équipements' }, { label: 'Dernier remplacement', nowrap: true }],
      res.parts.map((p) => [p.ref ? `<span class="tag">${esc(p.ref)}</span>` : '', esc(p.label), fmtN(p.qty), fmtN(p.cost), esc(p.eqs.join(', ')), fmtD(p.last)]),
      'Aucune pièce de rechange saisie sur la période.'
    );

    const prevRows = res.prevLate.map((t) => [t, 'late']).concat(res.prevSoon.map((t) => [t, 'soon']));
    $('prevSub').textContent = `${res.prevLate.length} en retard · ${res.prevSoon.length} à venir`;
    $('prevTable').innerHTML = table(
      [{ label: 'Échéance', nowrap: true }, { label: 'Équipement' }, { label: 'Opération' }, { label: 'Responsable' }, { label: 'Statut' }],
      prevRows.map(([t, s]) => [fmtD(t.due), eqName(t.code), esc(t.op), esc(t.resp), s === 'late' ? pill('bad', `En retard (${res.ref - t.due} j)`) : pill('warn', `Dans ${t.due - res.ref} j`)]),
      'Aucune tâche préventive en retard ni à venir sous 30 jours.'
    );

    $('otSub').textContent = `${res.otOpen.length} ouvert(s)`;
    $('otTable').innerHTML = table(
      [{ label: 'N° OT' }, { label: 'Équipement' }, { label: 'Description' }, { label: 'Priorité' }, { label: 'Date prévue', nowrap: true }, { label: 'Statut' }],
      res.otOpen.map((o) => [esc(o.id), eqName(o.code), esc(o.desc), esc(o.prio), fmtD(o.planned),
        o.planned != null && o.planned < res.ref ? pill('bad', `${o.status || 'Ouvert'} · en retard`) : `<span class="tag">${esc(o.status || 'Ouvert')}</span>`]),
      'Aucun ordre de travail ouvert.'
    );

    const P = res.recentPannes;
    const hasAlarm = P.some((p) => p.alarm), hasPart = P.some((p) => p.part);
    const pcols = [{ label: 'Date', nowrap: true }, { label: 'Équipement' }, { label: 'Sous-ensemble' }, { label: 'Description' }, { label: 'Cause' }];
    if (hasAlarm) pcols.push({ label: 'Code alarme' });
    if (hasPart) pcols.push({ label: 'Réf. composant' });
    pcols.push({ label: 'Arrêt (h)', num: true }, { label: 'Réparation (h)', num: true }, { label: 'Responsable' });
    const LIMIT = 300;
    $('pannesSub').textContent = `${P.length} panne(s) sur la période${P.length > LIMIT ? ` · ${LIMIT} plus récentes affichées` : ''}`;
    $('pannesTable').innerHTML = table(
      pcols,
      P.slice(0, LIMIT).map((p) => {
        const r = [fmtD(p.date), eqName(p.code), esc(p.family), esc(p.desc), esc(p.cause)];
        if (hasAlarm) r.push(esc(p.alarm));
        if (hasPart) r.push(p.part ? `<span class="tag">${esc(p.part)}</span>` : '');
        r.push(fmtN(p.downtime, 2), fmtN(p.repair, 2), `${esc(p.resp)}${p.include ? '' : ' <span class="tag">exclue des KPI</span>'}`);
        return r;
      }),
      'Aucune panne sur la période.'
    );

    const R = C.LEON_REFERENCE;
    $('refSource').textContent = `Source : ${R.source}`;
    $('refIntervals').innerHTML = table([{ label: 'Périodicité', nowrap: true }, { label: 'Opérations' }, { label: 'Par' }], R.intervals.map((r) => r.map(esc)));
    $('refConsumables').innerHTML = table([{ label: 'Pièce / consommable' }, { label: 'Remplacement' }], R.consumables.map((r) => r.map(esc)));
  }

  function renderDiag() {
    const rep = state.report;
    if (!rep) {
      $('diagBody').innerHTML = '<ul><li>Données de démonstration générées dans le navigateur.</li></ul>';
      return;
    }
    const names = { equipements: 'Équipements', pannes: 'Registre des pannes', interventions: 'Interventions', preventif: 'Plan préventif', ot: 'Ordres de travail' };
    $('diagBody').innerHTML = `<ul>${Object.entries(rep.tables)
      .map(([k, t]) =>
        t
          ? `<li><strong>${names[k]}</strong> : feuille « ${esc(t.sheet)} », en-têtes ligne ${t.headerRow}, ${t.rows} ligne(s).${t.missing.length ? ` Colonnes absentes (facultatives) : ${esc(t.missing.join(', '))}.` : ''}</li>`
          : `<li><strong>${names[k]}</strong> : non trouvé.</li>`
      )
      .join('')}<li>Paramètres : devise ${esc(state.data.params.devise)}, ${state.data.params.joursSem} jours d'exploitation / semaine${state.data.params.refDate ? `, date de référence ${fmtD(state.data.params.refDate)}` : ''}.</li></ul>`;
  }

  // ---------- wiring ----------
  function init() {
    const logo = $('logo');
    const markMissing = () => $('logoSlot').classList.add('missing');
    logo.addEventListener('error', markMissing);
    if (logo.complete && !logo.naturalWidth) markMissing();

    try {
      const t = localStorage.getItem('gm-theme');
      if (t) document.documentElement.dataset.theme = t;
    } catch (e) { /* storage unavailable */ }
    $('btnTheme').onclick = () => {
      const dark = document.documentElement.dataset.theme
        ? document.documentElement.dataset.theme === 'dark'
        : matchMedia('(prefers-color-scheme: dark)').matches;
      document.documentElement.dataset.theme = dark ? 'light' : 'dark';
      try { localStorage.setItem('gm-theme', document.documentElement.dataset.theme); } catch (e) { /* ignore */ }
      if (state.res) renderCharts(state.res);
    };
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => state.res && renderCharts(state.res));

    $('file').onchange = (e) => e.target.files[0] && loadFile(e.target.files[0]);
    $('btnDemo').onclick = loadDemo;
    $('btnTemplate').onclick = () => XLSX.writeFile(C.buildWorkbook(XLSX, null), 'modele_maintenance_leon.xlsx');
    $('btnExport').onclick = () => XLSX.writeFile(C.buildWorkbook(XLSX, state.data), state.isDemo ? 'demo_maintenance_leon.xlsx' : 'export_maintenance.xlsx');
    $('btnPrint').onclick = () => window.print();
    $('btnReset').onclick = () => {
      ['fEq', 'fModel', 'fZone', 'fCrit'].forEach((id) => ($(id).value = ''));
      $('fType').value = 'Tous';
      applyQuick('all');
    };
    ['fEq', 'fModel', 'fZone', 'fCrit', 'fType'].forEach((id) => $(id).addEventListener('change', render));
    ['fFrom', 'fTo'].forEach((id) => $(id).addEventListener('change', () => {
      document.querySelectorAll('#quick button').forEach((b) => b.setAttribute('aria-pressed', 'false'));
      render();
    }));
    document.querySelectorAll('#quick button').forEach((b) => (b.onclick = () => applyQuick(b.dataset.q)));

    let depth = 0;
    window.addEventListener('dragenter', (e) => { e.preventDefault(); depth++; $('drop').classList.add('on'); });
    window.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; $('drop').classList.remove('on'); } });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault();
      depth = 0;
      $('drop').classList.remove('on');
      const f = e.dataTransfer.files[0];
      if (f) loadFile(f);
    });

    loadDemo();
  }

  init();
  window.GMApp = { loadBuffer };
})();
