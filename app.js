/* Reporting Service Investissement — Euryale
 * Connexion Microsoft (MSAL) puis lecture des fichiers Excel SharePoint via Microsoft Graph,
 * avec les droits de la personne connectée. Aucune donnée n'est stockée dans la page.
 */
(() => {
  'use strict';

  const CFG = window.EURYALE_CONFIG || {};
  const SCOPES = ['Files.Read.All'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];

  /* ===================== Formats ===================== */
  const loc = d3.formatLocale({decimal: ',', thousands: ' ', grouping: [3], currency: ['', ' €']});
  const n0 = loc.format(',.0f');
  const n1 = loc.format(',.1f');
  const fmt = {
    int: (v) => (v == null ? '—' : n0(v)),
    meur: (v) => (v == null ? '—' : n1(v / 1e6) + ' M€'),
    smeur: (v) => (v == null ? '—' : (v > 0 ? '+' : '') + n1(v / 1e6) + ' M€'),
    eur: (v) => (v == null ? '—' : n0(v) + ' €'),
    pct: (v) => (v == null ? '—' : loc.format('.2f')(v * 100) + ' %'),
    pct1: (v) => (v == null ? '—' : loc.format('.1f')(v * 100) + ' %'),
    spct1: (v) => (v == null ? '—' : (v > 0 ? '+' : '') + loc.format('.1f')(v * 100) + ' %'),
    yrs: (v) => (v == null ? '—' : loc.format('.1f')(v) + ' ans'),
    date: (s) => { if (!s) return '—'; const p = String(s).split('-'); return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : s; },
    txt: (v) => (v == null || v === '' ? '—' : String(v)),
  };
  const PAYS = {France: 'France', Germany: 'Allemagne', Allemagne: 'Allemagne', UK: 'Royaume-Uni', 'United Kingdom': 'Royaume-Uni', 'Royaume-Uni': 'Royaume-Uni', Spain: 'Espagne', Espagne: 'Espagne', Italy: 'Italie', Italie: 'Italie', Canada: 'Canada', Netherlands: 'Pays-Bas', 'Pays-Bas': 'Pays-Bas', Ireland: 'Irlande', Irlande: 'Irlande', Portugal: 'Portugal', Belgium: 'Belgique', Belgique: 'Belgique'};
  const paysFr = (p) => PAYS[p] || p || '—';
  const C = {gold: '#c4993a', blue: '#4a6fa5', teal: '#3a8f8c', ok: '#2d9d6e', muted: '#c5cfdb'};

  /* ===================== Écran d'accueil / erreurs ===================== */
  const gate = (title, message, withButton) => {
    $('#gate').hidden = false;
    $('#main').hidden = true;
    $('#gate-t').textContent = title;
    $('#gate-m').textContent = message;
    $('#gate-btn').hidden = !withButton;
  };
  const showMain = () => { $('#gate').hidden = true; $('#main').hidden = false; };

  /* ===================== Connexion Microsoft ===================== */
  if (!CFG.clientId || /^A_REMPLIR/.test(CFG.clientId)) {
    gate('Site pas encore configuré', "L'identifiant d'application Microsoft n'est pas renseigné dans config.js. Voir la fiche d'installation transmise au service informatique.", false);
    return;
  }
  const here = new URL('./', location.href).href;
  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: `https://login.microsoftonline.com/${CFG.tenant || 'organizations'}`,
      redirectUri: new URL('redirect.html', here).href,
      postLogoutRedirectUri: here,
    },
    cache: {cacheLocation: 'localStorage'},
    system: {iframeBridgeTimeout: 8000},
  });

  const signIn = () => pca.loginRedirect({scopes: SCOPES, prompt: 'select_account'});
  $('#gate-btn').addEventListener('click', signIn);
  $('#btn-logout').addEventListener('click', () => pca.logoutRedirect({account: pca.getActiveAccount()}));
  $('#btn-refresh').addEventListener('click', () => loadAll());

  async function token() {
    const account = pca.getActiveAccount();
    try {
      return (await pca.acquireTokenSilent({scopes: SCOPES, account})).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        await pca.acquireTokenRedirect({scopes: SCOPES, account});
        return new Promise(() => {});
      }
      throw e;
    }
  }

  /* ===================== Lecture Excel via Microsoft Graph ===================== */
  class GraphError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
  const fileBase = (path) => `${GRAPH}/drives/${CFG.driveId}/root:/${path.split('/').map(encodeURIComponent).join('/')}:/workbook`;
  async function gget(url) {
    const r = await fetch(url, {headers: {Authorization: `Bearer ${await token()}`}});
    if (!r.ok) {
      let msg = '';
      try { msg = (await r.json()).error.message; } catch (_) { /* rien */ }
      throw new GraphError(r.status, msg || r.statusText);
    }
    return r.json();
  }
  async function readTable(path, name) {
    const base = `${fileBase(path)}/tables/${encodeURIComponent(name)}`;
    const [h, b] = await Promise.all([gget(`${base}/headerRowRange?$select=values`), gget(`${base}/dataBodyRange?$select=values`)]);
    return {head: h.values[0].map((x) => String(x).trim().toUpperCase()), rows: b.values};
  }
  async function readSheet(path, sheet) {
    const r = await gget(`${fileBase(path)}/worksheets/${encodeURIComponent(sheet)}/usedRange(valuesOnly=true)?$select=values`);
    return r.values;
  }

  // Conversion d'une cellule Excel (valeur brute renvoyée par Graph)
  const serialToIso = (n) => new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000).toISOString().slice(0, 10);
  function cell(v, type) {
    if (type === 't') return v == null || v === '' ? null : String(v).trim();
    if (type === 'd') {
      if (typeof v === 'number' && v > 0) return serialToIso(v);
      const m = String(v || '').trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
    }
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    let t = String(v || '').replace(/[\s  ]/g, '');
    if (!t) return null;
    const pct = t.endsWith('%');
    t = t.replace(/ans$/, '').replace(/m²$/, '').replace(/[€£$%]/g, '').replace(',', '.');
    if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
    return pct ? Number(t) / 100 : Number(t);
  }
  function rowsFrom(table, cols, keep) {
    const idx = cols.map(([h]) => table.head.indexOf(h));
    return table.rows.map((r) => {
      const o = {};
      cols.forEach(([, k, type], j) => { o[k] = idx[j] < 0 ? null : cell(r[idx[j]], type); });
      return o;
    }).filter(keep);
  }

  const COLS = {
    dealflow: [
      ['CODE', 'code', 't'], ['DATE DE RÉCÉPTION DU DOSSIER', 'reception', 'd'], ['TRIMESTRE', 'trimestre', 't'],
      ['STATUT DU DEAL', 'statut', 't'], ['NOM PORTEFEUILLE', 'portefeuille', 't'], ['NOM ACTIF', 'actif', 't'],
      ['GROUPE/EXPLOITANT', 'groupe', 't'], ['NOM BROKER', 'broker', 't'], ['PAYS', 'pays', 't'], ['VILLE', 'ville', 't'],
      ['TYPOLOGIE', 'typologie', 't'], ['SECTEUR', 'secteur', 't'], ['DEVISE', 'devise', 't'], ['LOYER ANNUEL', 'loyer', 'n'],
      ['PRIX AEM', 'prix_aem', 'n'], ['RENDEMENT BRUT AEM NORMATIF', 'rdt_aem', 'n'],
      ['RESPONSABLE INVEST (INITIALES)', 'responsable', 't'], ['COMMENTAIRES', 'commentaires', 't'],
    ],
    engagements: [
      ['FONDS', 'fonds', 't'], ['CODE ACTIF', 'code', 't'], ['NOM ACTIF', 'actif', 't'], ['GROUPE/EXPLOITANT', 'groupe', 't'],
      ['PAYS', 'pays', 't'], ['WALB', 'walb', 'n'], ['LOYER €/AN/HT/HC', 'loyer', 'n'], ['PRIX ACQUISITION AEM', 'prix_aem', 'n'],
      ['APPELS DE FONDS PASSES TTC', 'af_passes', 'n'], ['APPELS DE FONDS FUTURS TTC', 'af_futurs', 'n'],
      ['RENDEMENT BRUT AEM (AVEC PREMIUM)', 'rdt_aem', 'n'], ['COMPLETION %', 'completion', 'n'], ['DATE LIVRAISON PREVUE', 'livraison', 'd'],
    ],
    acquisitions: [
      ['FONDS', 'fonds', 't'], ['CODE ACTIF', 'code', 't'], ['NOM ACTIF', 'actif', 't'], ['GROUPE/EXPLOITANT', 'groupe', 't'],
      ["DATE D'ACQUISITON EURYALE", 'date_acq', 'd'], ['TRIMESTRE', 'trimestre', 't'], ['PAYS', 'pays', 't'], ['TYPOLOGIE', 'typologie', 't'],
      ['WALB', 'walb', 'n'], ['LOYER €/AN/HT/HC', 'loyer', 'n'], ['PRIX ACQUISITION AEM', 'prix_aem', 'n'], ['RENDEMENT BRUT AEM (AVEC PREMIUM)', 'rdt_aem', 'n'],
    ],
    cessions: [
      ['FONDS', 'fonds', 't'], ['CODE ACTIF', 'code', 't'], ['PORTEFEUILLE', 'portefeuille', 't'], ['PAYS', 'pays', 't'], ['VILLE', 'ville', 't'],
      ['DATE CESSION EFFECTIVE', 'date_cession', 'd'], ['DUREE DE DETENTION', 'detention', 'n'], ['PRIX ACQUISITION HD', 'prix_hd', 'n'],
      ['PRIX DE CESSION NET (TOUT FRAIS INCLUS)', 'prix_net', 'n'], ['PLUS-VALUE IMPOSABLE', 'pv', 'n'], ['PLUS OU MOINS VALUE %', 'pv_pct', 'n'],
    ],
  };
  const isMvRow = (r) => r.code && r.code.toLowerCase() !== 'total';

  function readTaux(grid) {
    for (let i = 0; i < grid.length; i++) {
      const iT = grid[i].findIndex((v) => String(v).trim() === 'Taux → EUR');
      if (iT < 0) continue;
      const iD = grid[i].findIndex((v) => String(v).trim() === 'Devise');
      const out = [];
      for (let j = i + 1; j < grid.length; j++) {
        const dv = String(grid[j][iD] ?? '').trim();
        if (!dv) break;
        out.push({devise: dv, taux: cell(grid[j][iT], 'n')});
      }
      return out;
    }
    throw new Error("Tableau « Devise / Taux → EUR » introuvable dans l'onglet Résumé");
  }

  /* ===================== Calculs (identiques au reporting Claude) ===================== */
  const LIB = {STUDIED: "À l'étude", 'EXPRESSION OF INTEREST': "Lettre d'intérêt", NEGOCIATIONS: 'Négociations', LOI: 'LOI', DD: 'Due diligence', INVESTED: 'Investi', 'REFUSED AFTER NEGO': 'Refusé après négo', REFUSED: 'Refusé', LOST: 'Perdu'};
  const ACTIFS = ['STUDIED', 'EXPRESSION OF INTEREST', 'NEGOCIATIONS', 'LOI', 'DD'];
  const ORDRE = ['STUDIED', 'EXPRESSION OF INTEREST', 'NEGOCIATIONS', 'LOI', 'DD', 'INVESTED', 'REFUSED AFTER NEGO', 'REFUSED', 'LOST'];
  const sum = (a, f) => a.reduce((s, d) => s + (d[f] || 0), 0);
  const wavg = (a, f, w) => { const tw = sum(a, w); return tw ? a.reduce((s, d) => s + (d[f] || 0) * (d[w] || 0), 0) / tw : null; };

  function computeDeals(dealflow, taux) {
    const rate = (dv) => {
      if (!dv) return 1;
      const t = taux.find((x) => x.devise === dv) || taux.find((x) => x.devise.startsWith(dv + ' '));
      return t && t.taux != null ? t.taux : null;
    };
    return dealflow.filter((d) => /^OP26/i.test(d.code || '')).map((d) => {
      const k = rate(d.devise);
      const statut = d.statut ? d.statut.toUpperCase() : null;
      return {
        ...d, statut, statut_libelle: LIB[statut] || d.statut || 'Non renseigné',
        etape: ACTIFS.includes(statut) ? 'En cours' : statut === 'INVESTED' ? 'Investi' : 'Clos',
        pays: PAYS[d.pays] || d.pays || 'Non renseigné',
        volume: d.prix_aem != null && k != null ? d.prix_aem * k : null,
        loyer: d.loyer != null && k != null ? d.loyer * k : null,
      };
    });
  }
  const groupBy = (a, keyFn, init, add) => {
    const m = new Map();
    a.forEach((d) => { const k = keyFn(d); const r = m.get(k) || init(k, d); add(r, d); m.set(k, r); });
    return [...m.values()];
  };
  function computeAll(raw) {
    const deals = computeDeals(raw.dealflow, raw.taux);
    const inv = deals.filter((d) => d.etape === 'Investi');
    const enc = deals.filter((d) => d.etape === 'En cours');
    const vol = sum(deals, 'volume');
    const loy = sum(deals, 'loyer');
    const dates = deals.map((d) => d.reception).filter(Boolean).sort();
    const dfKpi = {
      dossiers: deals.length, volume: vol, loyer: loy, rendement: vol ? loy / vol : null,
      investis: inv.length, volume_investi: sum(inv, 'volume'), en_cours: enc.length, volume_en_cours: sum(enc, 'volume'),
      refuses: deals.filter((d) => /^REFUSED/.test(d.statut || '')).length, perdus: deals.filter((d) => d.statut === 'LOST').length,
      taux_transfo: deals.length ? inv.length / deals.length : null, sans_devise: deals.filter((d) => !d.devise).length,
      dernier_dossier: dates.length ? dates[dates.length - 1] : null,
    };
    const rank = (k) => (ORDRE.indexOf(k) < 0 ? 99 : ORDRE.indexOf(k));
    const dfStatut = groupBy(deals, (d) => d.statut || 'NON RENSEIGNÉ', (k, d) => ({statut: k, libelle: d.statut_libelle, etape: d.etape, nb: 0, volume: 0}),
      (r, d) => { r.nb += 1; r.volume += d.volume || 0; }).sort((a, b) => rank(a.statut) - rank(b.statut));
    const dfTrim = groupBy(deals, (d) => d.trimestre || 'N.R.', (k) => ({trimestre: k, nb: 0, investis: 0, en_cours: 0, clos: 0, volume: 0}),
      (r, d) => { r.nb += 1; r.volume += d.volume || 0; if (d.etape === 'Investi') r.investis += 1; else if (d.etape === 'En cours') r.en_cours += 1; else r.clos += 1; })
      .sort((a, b) => String(a.trimestre).localeCompare(String(b.trimestre)));
    const dfPays = groupBy(deals, (d) => d.pays, (k) => ({pays: k, nb: 0, volume: 0, loyer: 0}),
      (r, d) => { r.nb += 1; r.volume += d.volume || 0; r.loyer += d.loyer || 0; })
      .map((r) => ({...r, rendement: r.volume ? r.loyer / r.volume : null})).sort((a, b) => b.volume - a.volume);
    const typos = groupBy(deals, (d) => (d.typologie || 'Non renseigné').trim().toUpperCase().replace(/^BUREAU$/, 'BUREAUX'),
      (k) => ({typologie: k, nb: 0, volume: 0}), (r, d) => { r.nb += 1; r.volume += d.volume || 0; }).sort((a, b) => b.nb - a.nb);
    const dfTypo = typos.slice(0, 8);
    if (typos.length > 8) dfTypo.push({typologie: 'AUTRES', nb: sum(typos.slice(8), 'nb'), volume: sum(typos.slice(8), 'volume')});
    dfTypo.forEach((r) => { r.libelle = r.typologie.length <= 5 ? r.typologie : r.typologie.charAt(0) + r.typologie.slice(1).toLowerCase(); });

    const A = raw.acquisitions;
    const E = raw.engagements;
    const S = raw.cessions;
    const acqKpi = {nb: A.length, prix_aem: sum(A, 'prix_aem'), loyer: sum(A, 'loyer'), rendement: wavg(A, 'rdt_aem', 'prix_aem'), walb: wavg(A, 'walb', 'prix_aem')};
    const engKpi = {nb: E.length, prix_aem: sum(E, 'prix_aem'), af_passes: sum(E, 'af_passes'), af_futurs: sum(E, 'af_futurs'), rendement: wavg(E, 'rdt_aem', 'prix_aem')};
    const cesKpi = {nb: S.length, prix_net: sum(S, 'prix_net'), pv: sum(S, 'pv'), pv_pct: wavg(S, 'pv_pct', 'prix_hd')};
    const tkey = (t) => { const x = String(t).match(/^(\d)T(\d{2})$/); return x ? x[2] + x[1] : 'zz' + t; };
    const acqTrim = groupBy(A, (a) => a.trimestre || 'N.R.', (k) => ({trimestre: k, PS: 0, EHS: 0, autre: 0, total: 0, nb: 0}),
      (r, a) => { const f = a.fonds === 'PS' || a.fonds === 'EHS' ? a.fonds : 'autre'; r[f] += a.prix_aem || 0; r.total += a.prix_aem || 0; r.nb += 1; })
      .sort((a, b) => tkey(a.trimestre).localeCompare(tkey(b.trimestre)));

    return {deals, dfKpi, dfStatut, dfTrim, dfPays, dfTypo, acqKpi, engKpi, cesKpi, acqTrim, acquisitions: A, engagements: E, cessions: S};
  }

  /* ===================== Graphiques ===================== */
  const textW = (svg, s, cls) => { const t = svg.append('text').attr('class', cls).text(s); const w = t.node().getComputedTextLength() || String(s).length * 7; t.remove(); return w; };
  function tip(el, sel, f) {
    const box = document.createElement('div');
    box.className = 'tip';
    box.hidden = true;
    el.append(box);
    sel.on('pointermove', (ev, r) => {
      box.replaceChildren(...f(r).map((t, i) => { const d = document.createElement('div'); if (i === 0) d.className = 'tip-t'; d.textContent = t; return d; }));
      box.hidden = false;
      const b = el.getBoundingClientRect();
      box.style.left = Math.max(0, Math.min(ev.clientX - b.left + 14, b.width - box.offsetWidth - 4)) + 'px';
      box.style.top = (ev.clientY - b.top + 14) + 'px';
    }).on('pointerleave', () => { box.hidden = true; });
  }
  function hbar(el, rows, o) {
    el.replaceChildren();
    if (!rows.length) { el.append(Object.assign(document.createElement('p'), {className: 'empty', textContent: 'Aucune donnée.'})); return; }
    const W = Math.max(240, el.clientWidth);
    const rowH = 32;
    const svg = d3.select(el).append('svg').attr('width', W).attr('height', rows.length * rowH + 8).attr('role', 'img').attr('aria-label', o.aria);
    const lw = Math.min(W * 0.42, d3.max(rows, (r) => textW(svg, o.label(r), 'lbl')) + 12);
    const vw = d3.max(rows, (r) => textW(svg, o.fmt(o.value(r)), 'val')) + 10;
    const x = d3.scaleLinear().domain([0, d3.max(rows, (r) => o.value(r) || 0) || 1]).range([0, Math.max(10, W - lw - vw)]);
    const g = svg.selectAll('g.row').data(rows).join('g').attr('class', 'row').attr('transform', (r, i) => `translate(0,${4 + i * rowH})`);
    g.append('rect').attr('width', W).attr('height', rowH).style('fill', 'transparent');
    g.append('text').attr('class', 'lbl').attr('x', lw - 12).attr('y', rowH / 2).attr('dy', '0.35em').attr('text-anchor', 'end').text((r) => o.label(r));
    g.append('rect').attr('x', lw).attr('y', 7).attr('height', rowH - 14).attr('rx', 4).attr('width', (r) => Math.max(2, x(o.value(r) || 0))).style('fill', (r) => o.color(r));
    g.append('text').attr('class', 'val').attr('x', (r) => lw + x(o.value(r) || 0) + 6).attr('y', rowH / 2).attr('dy', '0.35em').text((r) => o.fmt(o.value(r)));
    tip(el, g, o.tip);
  }
  function vbar(el, rows, o) {
    el.replaceChildren();
    if (!rows.length) { el.append(Object.assign(document.createElement('p'), {className: 'empty', textContent: o.empty || 'Aucune donnée.'})); return; }
    const lg = document.createElement('div');
    lg.className = 'legend';
    o.series.forEach((s) => { const it = document.createElement('span'); const sw = document.createElement('i'); sw.style.background = s.color; it.append(sw, document.createTextNode(s.label)); lg.append(it); });
    el.append(lg);
    const W = Math.max(240, el.clientWidth);
    const H = 230;
    const m = {t: 22, r: 4, b: 28, l: 4};
    const svg = d3.select(el).append('svg').attr('width', W).attr('height', H).attr('role', 'img').attr('aria-label', o.aria);
    const keys = o.series.map((s) => s.key);
    const x = d3.scaleBand().domain(rows.map(o.x)).range([m.l, W - m.r]).padding(0.32);
    const total = (r) => d3.sum(keys, (k) => r[k] || 0);
    const y = d3.scaleLinear().domain([0, d3.max(rows, total) || 1]).nice().range([H - m.b, m.t]);
    svg.append('line').attr('x1', m.l).attr('x2', W - m.r).attr('y1', y(0)).attr('y2', y(0)).style('stroke', '#c9d3df');
    d3.stack().keys(keys)(rows).forEach((layer, li) => {
      svg.append('g').selectAll('rect').data(layer).join('rect').attr('x', (d) => x(o.x(d.data))).attr('width', x.bandwidth())
        .attr('y', (d) => y(d[1])).attr('height', (d) => Math.max(0, y(d[0]) - y(d[1]))).style('fill', o.series[li].color);
    });
    svg.append('g').selectAll('text').data(rows).join('text').attr('class', 'val').attr('x', (r) => x(o.x(r)) + x.bandwidth() / 2).attr('y', (r) => y(total(r)) - 6).attr('text-anchor', 'middle').text((r) => o.fmt(total(r)));
    svg.append('g').selectAll('text').data(rows).join('text').attr('class', 'lbl').attr('x', (r) => x(o.x(r)) + x.bandwidth() / 2).attr('y', H - 8).attr('text-anchor', 'middle').text((r) => o.x(r));
    const hit = svg.append('g').selectAll('rect').data(rows).join('rect').attr('x', (r) => x(o.x(r)) - (x.step() - x.bandwidth()) / 2).attr('width', x.step()).attr('y', m.t).attr('height', H - m.t - m.b).style('fill', 'transparent');
    tip(el, hit, o.tip);
  }
  function tbody(el, rows, cols, empty) {
    if (!rows.length) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = cols.length; td.className = 'empty'; td.textContent = empty || 'Aucune ligne.';
      tr.append(td); el.replaceChildren(tr); return;
    }
    el.replaceChildren(...rows.map((r) => {
      const tr = document.createElement('tr');
      cols.forEach(([f, ff, cls]) => {
        const td = document.createElement('td');
        if (cls) td.className = cls;
        td.textContent = ff(r[f], r);
        if (cls && cls.includes('signed') && typeof r[f] === 'number' && r[f] < 0) td.classList.add('is-bad');
        tr.append(td);
      });
      return tr;
    }));
  }

  /* ===================== Affichage ===================== */
  let D = null;
  const state = {statut: '', pays: '', q: ''};
  const plural = (n, s) => `${fmt.int(n)} ${s}${n > 1 ? 's' : ''}`;

  function fillNumbers() {
    $$('[data-k]').forEach((el) => {
      const [src, field] = el.getAttribute('data-k').split('.');
      const v = D && D[src] ? D[src][field] : null;
      el.textContent = (fmt[el.getAttribute('data-f')] || fmt.txt)(v);
      el.classList.remove('sk');
      if (el.getAttribute('data-tone') === 'sign') { el.classList.toggle('is-bad', v != null && v < 0); el.classList.toggle('is-good', v != null && v > 0); }
    });
  }
  function drawCharts() {
    if (!D) return;
    const vis = (id) => { const el = document.getElementById(id); return el && el.offsetParent !== null ? el : null; };
    let el;
    if ((el = vis('ch-statut'))) {
      hbar(el, D.dfStatut, {label: (r) => r.libelle, value: (r) => r.nb, fmt: fmt.int, aria: 'Nombre de dossiers 2026 par statut',
        color: (r) => (r.etape === 'Investi' ? C.ok : r.etape === 'En cours' ? C.gold : C.muted),
        tip: (r) => [r.libelle, plural(r.nb, 'dossier'), `${fmt.meur(r.volume)} de volume`]});
    }
    if ((el = vis('ch-trim'))) {
      vbar(el, D.dfTrim, {x: (r) => r.trimestre, fmt: fmt.int, aria: 'Dossiers reçus par trimestre, par étape',
        series: [{key: 'investis', label: 'Investis', color: C.ok}, {key: 'en_cours', label: 'En cours', color: C.gold}, {key: 'clos', label: 'Refusés ou perdus', color: C.muted}],
        tip: (r) => [r.trimestre, `${plural(r.nb, 'dossier')} reçus`, `${fmt.int(r.investis)} investis · ${fmt.int(r.en_cours)} en cours · ${fmt.int(r.clos)} clos`, `${fmt.meur(r.volume)} de volume`]});
    }
    if ((el = vis('ch-pays'))) {
      hbar(el, D.dfPays, {label: (r) => r.pays, value: (r) => r.volume, fmt: fmt.meur, color: () => C.blue, aria: 'Volume étudié par pays',
        tip: (r) => [r.pays, `${fmt.meur(r.volume)} · ${plural(r.nb, 'dossier')}`, `Loyer ${fmt.meur(r.loyer)} · rendement ${fmt.pct(r.rendement)}`]});
    }
    if ((el = vis('ch-typo'))) {
      hbar(el, D.dfTypo, {label: (r) => r.libelle, value: (r) => r.nb, fmt: fmt.int, color: () => C.teal, aria: 'Nombre de dossiers 2026 par typologie',
        tip: (r) => [r.libelle, plural(r.nb, 'dossier'), `${fmt.meur(r.volume)} de volume`]});
    }
    if ((el = vis('ch-acq-trim'))) {
      const series = [{key: 'PS', label: 'Fonds PS', color: C.gold}, {key: 'EHS', label: 'Fonds EHS', color: C.blue}];
      if (D.acqTrim.some((r) => r.autre)) series.push({key: 'autre', label: 'Autre fonds', color: C.teal});
      vbar(el, D.acqTrim, {x: (r) => r.trimestre, fmt: fmt.meur, series, aria: "Prix d'acquisition AEM par trimestre et par fonds", empty: 'Aucune acquisition enregistrée.',
        tip: (r) => [r.trimestre, `${fmt.meur(r.total)} · ${plural(r.nb, 'ligne')}`, `PS ${fmt.meur(r.PS)} · EHS ${fmt.meur(r.EHS)}`]});
    }
  }
  const key = (v) => (v == null ? '' : String(v));
  const byDesc = (f) => (a, b) => key(b[f]).localeCompare(key(a[f]));
  const byAsc = (f) => (a, b) => (a[f] == null) - (b[f] == null) || key(a[f]).localeCompare(key(b[f]));
  const colsDeal = (last) => [
    ['code', fmt.txt, 'code'], ['reception', fmt.date], ['statut_libelle', fmt.txt, 'pill'], ['actif', fmt.txt, 'strong'],
    ['groupe', fmt.txt], ['pays', fmt.txt], ['typologie', fmt.txt], ['volume', fmt.eur, 'num'], ['rdt_aem', fmt.pct, 'num'], last,
  ];
  function filtered() {
    const q = state.q.trim().toLowerCase();
    return D.deals.filter((d) => (!state.statut || d.statut === state.statut) && (!state.pays || d.pays === state.pays)
      && (!q || [d.code, d.actif, d.portefeuille, d.groupe, d.ville, d.typologie, d.commentaires].some((v) => v && String(v).toLowerCase().includes(q))));
  }
  function drawListe() {
    const rows = filtered().sort(byDesc('reception'));
    tbody($('#tb-liste'), rows, colsDeal(['commentaires', fmt.txt, 'note']), 'Aucun dossier ne correspond aux filtres.');
    $('#f-n').textContent = fmt.int(rows.length);
  }
  function drawTables() {
    tbody($('#tb-pipeline'), D.deals.filter((d) => d.etape === 'En cours').sort(byDesc('reception')), colsDeal(['responsable', fmt.txt]), 'Aucun dossier en cours.');
    const fs = $('#f-statut');
    fs.replaceChildren(new Option('Tous les statuts', ''), ...D.dfStatut.map((r) => new Option(`${r.libelle} (${r.nb})`, r.statut)));
    fs.value = state.statut;
    const fp = $('#f-pays');
    fp.replaceChildren(new Option('Tous les pays', ''), ...[...D.dfPays].sort((a, b) => a.pays.localeCompare(b.pays, 'fr')).map((r) => new Option(`${r.pays} (${r.nb})`, r.pays)));
    fp.value = state.pays;
    drawListe();
    tbody($('#tb-acq'), [...D.acquisitions].sort(byAsc('date_acq')), [
      ['code', fmt.txt, 'code'], ['fonds', fmt.txt, 'pill'], ['actif', fmt.txt, 'strong'], ['groupe', fmt.txt], ['pays', paysFr],
      ['typologie', fmt.txt], ['date_acq', fmt.date], ['prix_aem', fmt.eur, 'num'], ['loyer', fmt.eur, 'num'], ['rdt_aem', fmt.pct, 'num'], ['walb', fmt.yrs, 'num'],
    ], 'Aucune acquisition.');
    tbody($('#tb-eng'), [...D.engagements].sort(byAsc('livraison')), [
      ['code', fmt.txt, 'code'], ['fonds', fmt.txt, 'pill'], ['actif', fmt.txt, 'strong'], ['groupe', fmt.txt], ['pays', paysFr],
      ['livraison', fmt.date], ['prix_aem', fmt.eur, 'num'], ['af_passes', fmt.eur, 'num'], ['af_futurs', fmt.eur, 'num'], ['completion', fmt.pct1, 'num'], ['rdt_aem', fmt.pct, 'num'],
    ], 'Aucun engagement en cours.');
    tbody($('#tb-ces'), [...D.cessions].sort(byAsc('date_cession')), [
      ['code', fmt.txt, 'code'], ['portefeuille', fmt.txt, 'strong'], ['ville', fmt.txt], ['pays', paysFr], ['date_cession', fmt.date],
      ['detention', fmt.yrs, 'num'], ['prix_net', fmt.eur, 'num'], ['pv', fmt.eur, 'num signed'], ['pv_pct', fmt.spct1, 'num signed'],
    ], 'Aucune cession.');
  }
  $('#f-statut').addEventListener('change', (e) => { state.statut = e.target.value; if (D) drawListe(); });
  $('#f-pays').addEventListener('change', (e) => { state.pays = e.target.value; if (D) drawListe(); });
  $('#f-q').addEventListener('input', (e) => { state.q = e.target.value; if (D) drawListe(); });

  /* ===================== Onglets, impression, redimensionnement ===================== */
  const tabs = $$('.tab[data-tab]');
  function openTab(id, push) {
    tabs.forEach((t) => {
      const on = t.getAttribute('data-tab') === id;
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      document.getElementById(t.getAttribute('data-tab')).hidden = !on;
    });
    if (push) history.replaceState(null, '', `#${id}`);
    drawCharts();
  }
  tabs.forEach((t) => t.addEventListener('click', () => openTab(t.getAttribute('data-tab'), true)));
  const h0 = location.hash.replace('#', '');
  if (tabs.some((t) => t.getAttribute('data-tab') === h0)) openTab(h0, false);
  window.addEventListener('beforeprint', () => { $$('.section').forEach((s) => { s.dataset.wasHidden = s.hidden ? '1' : ''; s.hidden = false; }); drawCharts(); });
  window.addEventListener('afterprint', () => { $$('.section').forEach((s) => { s.hidden = s.dataset.wasHidden === '1'; }); drawCharts(); });
  let rz = null;
  window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(drawCharts, 150); });

  /* ===================== Chargement ===================== */
  function explain(e) {
    if (e instanceof GraphError) {
      if (e.status === 401 || e.status === 403) return "Votre compte n'a pas accès à l'un des fichiers SharePoint du reporting (dossier « Reporting Service Investissement » du site Equipe Invest). Demandez l'accès à son propriétaire.";
      if (e.status === 404) return `Fichier ou tableau introuvable : il a peut-être été renommé, déplacé, ou le tableau Excel a changé de nom (voir config.js). Détail : ${e.message}`;
      return `Microsoft Graph a répondu « ${e.message} » (code ${e.status}).`;
    }
    return e && e.message ? e.message : String(e);
  }
  let loading = false;
  let lastLoad = 0;
  async function loadAll() {
    if (loading) return;
    loading = true;
    lastLoad = Date.now();
    $('#u-stamp').textContent = 'lecture des fichiers…';
    try {
      const df = CFG.dealflow;
      const mv = CFG.mouvements;
      const [tDeal, grid, tEng, tAcq, tCes] = await Promise.all([
        readTable(df.path, df.table), readSheet(df.path, df.resumeSheet),
        readTable(mv.path, mv.engagements), readTable(mv.path, mv.acquisitions), readTable(mv.path, mv.cessions),
      ]);
      D = computeAll({
        dealflow: rowsFrom(tDeal, COLS.dealflow, (r) => /^OP/i.test(r.code || '')),
        taux: readTaux(grid),
        engagements: rowsFrom(tEng, COLS.engagements, isMvRow),
        acquisitions: rowsFrom(tAcq, COLS.acquisitions, isMvRow),
        cessions: rowsFrom(tCes, COLS.cessions, isMvRow),
      });
      showMain();
      fillNumbers();
      drawTables();
      drawCharts();
      $('#u-stamp').textContent = `données lues à ${new Date().toLocaleTimeString('fr-FR', {hour: '2-digit', minute: '2-digit'})}`;
    } catch (e) {
      if (!D) gate('Lecture impossible', explain(e), false);
      else $('#u-stamp').textContent = `échec de l'actualisation : ${explain(e)}`;
    } finally {
      loading = false;
    }
  }

  async function start() {
    try {
      await pca.initialize();
      const res = await pca.handleRedirectPromise();
      let account = (res && res.account) || pca.getActiveAccount() || pca.getAllAccounts()[0];
      if (!account) {
        gate('Connexion en cours…', 'Connexion automatique avec votre compte Microsoft Euryale. Si rien ne se passe, cliquez sur le bouton.', true);
        try { account = (await pca.ssoSilent({scopes: SCOPES})).account; } catch (_) { /* connexion manuelle */ }
      }
      if (!account) {
        gate('Connectez-vous', 'Ce reporting lit les fichiers SharePoint de l’équipe Investissement. Connectez-vous avec votre compte Microsoft Euryale.', true);
        return;
      }
      pca.setActiveAccount(account);
      $('#u-name').textContent = account.name || account.username;
      $('#user').hidden = false;
      gate('Chargement…', 'Lecture des fichiers Dealflow et Mouvements patrimoine sur SharePoint.', false);
      await loadAll();
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible' && D && Date.now() - lastLoad > 15 * 60 * 1000) loadAll();
      });
    } catch (e) {
      gate('Connexion impossible', explain(e), true);
    }
  }
  start();
})();
