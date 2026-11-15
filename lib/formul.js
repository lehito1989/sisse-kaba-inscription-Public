/* SISSÊ~KABA — module « Formulation & rentabilité MG2MIX ».
   Le classeur SOLEVO (FORMULATION ALIMENTAIRE AVEC LES PREMIX SOLEVO) est évalué tel quel par le moteur SKX :
   chaque saisie du visiteur devient une valeur de la feuille REGISTRE TERRAIN, et toutes les formules du modèle
   (formulations, coût du poulet, coût de l'œuf, bâtiments, devis) sont recalculées. Fonctionne aussi sous Node (tests). */
(function (root) {
  'use strict';
  const RT = 'REGISTRE TERRAIN', SC = 'MG2MIX 2% CHAIRS MARCHE', SP = 'MG2MIX2% PONTE MARCHE', CPC = 'COÛT POULET CHAIR', CO = 'COÛT ŒUF',
    DC = 'DEVIS CHAIR', DP = 'DEVIS PONDEUSES', BC = 'BÂTIMENT CHAIR', BP = 'BÂTIMENT PONDEUSES';
  const EUR = 656;

  /* matières premières : clé → cellule du registre */
  const MP = [
    ['mais', 'Maïs', 'B5'], ['soja', 'Tourteau de soja', 'B6'], ['son', 'Son de blé', 'B7'], ['coton', 'Tourteau de coton', 'B8'],
    ['coq', 'Coquillages', 'B9'], ['huile', 'Huile de palme', 'B10'], ['pxChair', 'Prémix chair MG2MIX 2 %', 'B11'],
    ['pxPonte', 'Prémix ponte MG2MIX 2 %', 'B12'], ['poisson', 'Farine de poisson', 'B13']];
  /* paramètres d'élevage modifiables */
  const PAR = {
    chair: [['effectif', 'Poussins mis en place', 'B29', 'sujets'], ['poussin', 'Prix du poussin', 'B30', 'FCFA'], ['mortalite', 'Mortalité', 'B31', '%'],
      ['poids', 'Poids vif à la vente', 'B32', 'kg'], ['prixVente', 'Prix de vente du poulet', 'B17', 'FCFA'], ['bandes', 'Bandes par an', 'B110', 'bandes']],
    ponte: [['effectif', 'Poussins ponte mis en place', 'B50', 'sujets'], ['poussin', 'Prix du poussin ponte', 'B51', 'FCFA'], ['mortalite', 'Mortalité poulette', 'B52', '%'],
      ['mortP', 'Mortalité en ponte', 'B53', '%'], ['prixPlaq', 'Prix de vente de la plaquette', 'B16', 'FCFA'], ['prixRef', 'Prix de la poule de réforme', 'B18', 'FCFA']]
  };
  /* formules du classeur */
  const CHAIR_F = [
    { id: 1, nom: 'Maïs / soja / son', cols: ['B', 'C', 'D', 'E', 'F'], r1: 14, r2: 23, kg: 13, sac: 12 },
    { id: 2, nom: 'Concentré + tourteau de coton', cols: ['H', 'I', 'J', 'K', 'L'], r1: 14, r2: 23, kg: 13, sac: 12 },
    { id: 3, nom: 'Farine de poisson', cols: ['B', 'C', 'D', 'E', 'F'], r1: 31, r2: 40, kg: 30, sac: 29 },
    { id: 4, nom: 'Huile de palme', cols: ['H', 'I', 'J', 'K', 'L'], r1: 31, r2: 40, kg: 30, sac: 29 }];
  const CHAIR_PH = [['Démarrage', '0-10 j'], ['Croissance', '11-28 j'], ['Finition', '29-35 j']];
  const PONTE_F = [
    { id: 1, key: 'MAIS/SOJA/SON', nom: 'Maïs / soja / son', cols: ['B', 'C', 'E', 'F', 'G', 'H', 'I', 'J'], r1: 15, r2: 25, kg: 14, sac: 13 },
    { id: 2, key: 'HUILE DE PALME', nom: 'Huile de palme', cols: ['B', 'C', 'E', 'F', 'G', 'H', 'I', 'J'], r1: 35, r2: 43, kg: 34, sac: 33 },
    { id: 3, key: 'FARINE DE POISSON', nom: 'Farine de poisson', cols: ['L', 'M', 'O', 'P', 'Q', 'R', 'S', 'T'], r1: 15, r2: 25, kg: 14, sac: 13 },
    { id: 4, key: 'T. COTON', nom: 'Tourteau de coton', cols: ['L', 'M', 'O', 'P', 'Q', 'R', 'S', 'T'], r1: 35, r2: 43, kg: 34, sac: 33 }];
  const PONTE_PH = [['Poulette démarrage', '0-4 sem'], ['Poulette 1er âge', '5-11 sem'], ['Poulette 2e âge', '12-18 sem'],
    ['Entrée en ponte', '19-35 sem'], ['Milieu de ponte', '36-50 sem'], ['Fin de ponte', '> 50 sem']];

  let WB = null, BASE = null;
  function b64(s) {
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(s, 'base64'));
    const bin = atob(s), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u;
  }
  async function freshWb() { const JSZip = root.JSZip; const zip = await JSZip.loadAsync(b64(root.SKF_XLSX)); return root.SKX.loadWorkbook(zip); }
  async function load() {
    if (WB) return WB;
    WB = await freshWb();
    const v = root.SKX.evaluate(WB, {});
    const g = (s, r) => (WB.sheets[s].cells[r] || {}).value;
    BASE = { prix: {}, chair: {}, ponte: {} };
    MP.forEach(([k, , c]) => { BASE.prix[k] = g(RT, c); });
    ['chair', 'ponte'].forEach((sp) => PAR[sp].forEach(([k, , c]) => { BASE[sp][k] = g(RT, c); }));
    BASE.v = v;
    return WB;
  }
  const loaded = () => !!WB;

  /* saisie → valeurs du registre */
  function overrides(inp) {
    const o = {};
    const P = Object.assign({}, BASE.prix, inp.prix || {});
    MP.forEach(([k, , c]) => { if (isFinite(+P[k]) && +P[k] > 0) o[RT + '!' + c] = +P[k]; });
    ['chair', 'ponte'].forEach((sp) => {
      const S = Object.assign({}, BASE[sp], inp[sp] || {});
      PAR[sp].forEach(([k, , c, u]) => { let x = +S[k]; if (!isFinite(x)) return; if (u === '%' && x > 1) x = x / 100; o[RT + '!' + c] = x; });
    });
    if (inp.batiment === 'existant') { o[RT + '!B45'] = 0; o[RT + '!B75'] = 0; }
    o[CO + '!C17'] = (PONTE_F.find((f) => f.id === +(inp.ponte && inp.ponte.formule)) || PONTE_F[0]).key;
    const cl = inp.client || {};
    [[DC, 'C6', cl.nom], [DC, 'C7', cl.localite], [DC, 'C8', cl.tel], [DP, 'C6', cl.nom], [DP, 'C7', cl.localite], [DP, 'C8', cl.tel]]
      .forEach(([s, r, x]) => { if (x) o[s + '!' + r] = String(x); });
    return o;
  }
  /* la formule chair utilisée par COÛT POULET CHAIR (G18:G20) pointe vers le bloc choisi */
  function setChair(wb, id) {
    const f = CHAIR_F.find((x) => x.id === +id) || CHAIR_F[0];
    const c = wb.sheets[CPC].cells;
    [0, 1, 2].forEach((i) => { c['G' + (18 + i)].f = `'${SC}'!${f.cols[2 + i]}${f.kg}`; });
    return f;
  }
  function blockOf(v, sheet, f, nph) {
    const g = (c, r) => v[sheet + '!' + c + r];
    const rows = [];
    for (let r = f.r1; r <= f.r2; r++) {
      const nom = g(f.cols[0], r); if (!nom || typeof nom !== 'string') continue;
      const kg = []; for (let i = 0; i < nph; i++) { const x = g(f.cols[2 + i], r); kg.push(typeof x === 'number' ? x : 0); }
      if (!kg.some((x) => x > 0)) continue;
      rows.push({ nom: nom.trim(), prix: +g(f.cols[1], r) || 0, kg });
    }
    const kgP = [], sac = []; for (let i = 0; i < nph; i++) { kgP.push(+g(f.cols[2 + i], f.kg) || 0); sac.push(+g(f.cols[2 + i], f.sac) || 0); }
    return { id: f.id, nom: f.nom, rows, prixKg: kgP, prixSac: sac };
  }
  const N = (x) => (typeof x === 'number' && isFinite(x) ? x : 0);

  function evalAll(inp, chairId, ponteId) {
    const o = overrides(Object.assign({}, inp, { ponte: Object.assign({}, inp.ponte, { formule: ponteId }) }));
    setChair(WB, chairId);
    const v = root.SKX.evaluate(WB, o);
    setChair(WB, 1);
    return v;
  }
  function premixNeed(block, qty) {
    const out = { chair: 0, ponte: 0 };
    block.rows.forEach((r) => { const t = /PREMIX CHAIR/i.test(r.nom) ? 'chair' : /PREMIX PONTE/i.test(r.nom) ? 'ponte' : null; if (t) r.kg.forEach((k, i) => { out[t] += k * (qty[i] || 0) / 1000; }); });
    return out;
  }
  /* calcul complet */
  function compute(inp) {
    inp = inp || {};
    const cId = +((inp.chair || {}).formule) || 1, pId = +((inp.ponte || {}).formule) || 1;
    const v = evalAll(inp, cId, pId);
    const g = (s, r) => N(v[s + '!' + r]);
    const qC = [18, 19, 20].map((r) => g(CPC, 'J' + r)), qP = [21, 22, 23, 24, 25, 26].map((r) => g(CO, 'J' + r));
    const bC = blockOf(v, SC, CHAIR_F.find((f) => f.id === cId), 3), bP = blockOf(v, SP, PONTE_F.find((f) => f.id === pId), 6);
    const chair = {
      formule: bC, phases: CHAIR_PH.map((p, i) => ({ nom: p[0], age: p[1], jours: g(CPC, 'D' + (18 + i)), kgSujet: g(CPC, 'F' + (18 + i)), qte: qC[i], prixKg: g(CPC, 'H' + (18 + i)), cout: g(CPC, 'K' + (18 + i)) })),
      effectif: g(CPC, 'C6'), vendus: g(CPC, 'C14'), duree: g(CPC, 'C13'), aliment: g(CPC, 'J21'), coutAliment: g(CPC, 'C37'), horsAliment: g(CPC, 'C38'),
      coutTotal: g(CPC, 'C39'), partAliment: g(CPC, 'C40'), coutPoulet: g(CPC, 'C41'), coutKgVif: g(CPC, 'C42'), ic: g(CPC, 'C43'), prixVente: g(CPC, 'C44'),
      margePoulet: g(CPC, 'C46'), margeBande: g(CPC, 'C47'), margePct: g(CPC, 'C48'), seuil: g(CPC, 'C49'),
      surface: g(BC, 'C8'), batiment: g(DC, 'F22'), equipements: g(DC, 'F23'), invest: g(DC, 'F24'), chargesBande: g(DC, 'F38'), coutHorsAmort: g(DC, 'F39'),
      demarrage: g(DC, 'F44'), caAn: g(DC, 'F48'), chargesAn: g(DC, 'F49'), margeBruteAn: g(DC, 'F50'), amortAn: g(DC, 'F51'), margeAn: g(DC, 'F52'), mensuel: g(DC, 'F53'),
      retour: typeof v[DC + '!F54'] === 'number' ? v[DC + '!F54'] : null, bandes: g(BC, 'C48') || g(RT, 'B110'),
      premix: premixNeed(bC, qC), sacs: Math.ceil(g(CPC, 'J21') / (g(RT, 'B33') || 50))
    };
    const ponte = {
      formule: bP, phases: PONTE_PH.map((p, i) => ({ nom: p[0], age: p[1], jours: g(CO, 'D' + (21 + i)), kgSujet: g(CO, 'F' + (21 + i)), qte: qP[i], prixKg: g(CO, 'H' + (21 + i)), cout: g(CO, 'K' + (21 + i)) })),
      effectif: g(CO, 'C6'), poulettes: g(CO, 'C117'), duree: g(CO, 'C15'), aliment: g(CO, 'J27'), coutAliment: g(CO, 'C53'), horsAliment: g(CO, 'C54'), coutTotal: g(CO, 'C55'),
      partAliment: g(CO, 'C56'), coutOeuf: g(CO, 'C57'), coutPlaq: g(CO, 'C58'), prixPlaq: g(CO, 'C59'), margeOeuf: g(CO, 'C61'), margePlaq: g(CO, 'C62'),
      coutPoulette: g(CO, 'C118'), oeufs: g(CO, 'G35'), plaquettes: g(CO, 'G36'), surface: g(BP, 'C8'),
      batiment: g(DP, 'F24'), equipements: g(DP, 'F25'), invest: g(DP, 'F26'), chargesCycle: g(DP, 'F49'), demarrage: g(DP, 'F57'),
      caOeufsAn: g(DP, 'F61'), caReformeAn: g(DP, 'F62'), caAn: g(DP, 'F63'), chargesAn: g(DP, 'F64'), margeBruteAn: g(DP, 'F65'), amortAn: g(DP, 'F66'), margeAn: g(DP, 'F67'), mensuel: g(DP, 'F68'),
      retour: typeof v[DP + '!F69'] === 'number' ? v[DP + '!F69'] : null,
      premix: premixNeed(bP, qP), sacs: Math.ceil(g(CO, 'J27') / (g(RT, 'B55') || 50))
    };
    if (inp.batiment === 'existant') ['batiment', 'equipements', 'invest'].forEach((k) => { chair[k] = 0; ponte[k] = 0; });
    if (inp.batiment === 'existant') { chair.demarrage = chair.chargesBande; ponte.demarrage = g(DP, 'F56'); chair.retour = null; ponte.retour = null; }
    /* comparaison des 4 formules (mêmes hypothèses) */
    const comp = { chair: [], ponte: [] };
    for (let i = 1; i <= 4; i++) {
      const w = i === cId && i === pId ? v : evalAll(inp, i, i);
      const G = (s, r) => N(w[s + '!' + r]);
      comp.chair.push({ id: i, nom: CHAIR_F[i - 1].nom, prixMoyen: G(CPC, 'J21') ? G(CPC, 'C37') / G(CPC, 'J21') : 0, coutAliment: G(CPC, 'C37'), coutPoulet: G(CPC, 'C41'), marge: G(CPC, 'C46'), margeBande: G(CPC, 'C47') });
      const pf = PONTE_F.find((f) => f.id === i);
      comp.ponte.push({ id: i, nom: pf.nom, prixMoyen: G(CO, 'J27') ? G(CO, 'C53') / G(CO, 'J27') : 0, coutAliment: G(CO, 'C53'), coutPlaq: G(CO, 'C58'), marge: G(CO, 'C62'), margeCycle: G(CO, 'C62') * G(CO, 'G36') });
    }
    ['chair', 'ponte'].forEach((k) => { const m = Math.max(...comp[k].map((x) => x.marge)); comp[k].forEach((x) => { x.best = x.marge === m; }); });
    return { chair, ponte, comp, prix: Object.assign({}, BASE.prix, inp.prix || {}), base: BASE };
  }
  /* classeur personnalisé : registre rempli, formules conservées et recalculées */
  async function exportXlsx(inp) {
    const wb = await freshWb();
    const o = overrides(inp);
    setChair(wb, +((inp.chair || {}).formule) || 1);
    const v = root.SKX.evaluate(wb, o);
    root.SKX.writeBack(wb, v, o);
    const zip = wb.zip;
    /* recalcul à l'ouverture dans Excel */
    const wbx = await zip.file('xl/workbook.xml').async('string');
    if (!/fullCalcOnLoad/.test(wbx)) zip.file('xl/workbook.xml', wbx.replace(/<calcPr([^>]*)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>').replace(/<\/workbook>/, /<calcPr/.test(wbx) ? '</workbook>' : '<calcPr fullCalcOnLoad="1"/></workbook>'));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /* rapport visiteur (texte WhatsApp + HTML e-mail), commun à la fiche publique et au simulateur.
     o = { nom, localite, sps: ['chair','ponte'], valide: bool } */
  function fr(x, d) {
    if (typeof x !== 'number' || !isFinite(x)) return 'n.d.';
    const neg = x < 0; let [i, f] = Math.abs(x).toFixed(d || 0).split('.'); i = i.replace(/\B(?=(\d{3})+(?!\d))/g, '\u00a0');
    return (neg ? '−' : '') + i + (f ? ',' + f : '');
  }
  const eu = (x) => fr(x) + '\u00a0FCFA (≈\u00a0' + fr(x / EUR) + '\u00a0€)';
  function report(R, o) {
    o = o || {}; const sps = o.sps && o.sps.length ? o.sps : ['chair', 'ponte'];
    const sec = [];
    if (sps.includes('chair')) {
      const c = R.chair;
      sec.push({ t: '🐔 POULET DE CHAIR – ' + fr(c.effectif) + ' sujets, bande de ' + fr(c.duree) + ' jours', f: 'Aliment formulé ' + c.formule.nom + ' + MG2MIX 2 %',
        ph: c.phases.map((p, i) => [p.nom + ' (' + p.age + ')', fr(c.formule.prixKg[i], 1) + ' FCFA/kg • ' + fr(c.formule.prixSac[i]) + ' FCFA le sac de 50 kg']),
        kv: [['Coût de revient d’un poulet', fr(c.coutPoulet) + ' FCFA (' + fr(c.coutKgVif) + ' FCFA/kg vif)'], ['Prix de vente retenu', fr(c.prixVente) + ' FCFA/poulet'],
          ['Marge brute par poulet', fr(c.margePoulet) + ' FCFA (' + fr(c.margePct * 100) + ' %)'], ['Marge brute de la bande', eu(c.margeBande)],
          ['Chiffre d’affaires annuel (' + fr(c.bandes) + ' bandes)', eu(c.caAn)], ['Marge annuelle après amortissement (estimation)', eu(c.margeAn)],
          ['Prémix MG2MIX chair nécessaire', fr(c.premix.chair) + ' kg par bande']] });
    }
    if (sps.includes('ponte')) {
      const p = R.ponte;
      sec.push({ t: '🥚 PONDEUSES – ' + fr(p.effectif) + ' poussins, cycle de 78 semaines', f: 'Aliment formulé ' + p.formule.nom + ' + MG2MIX 2 %',
        ph: p.phases.map((x, i) => [x.nom + ' (' + x.age + ')', fr(p.formule.prixKg[i], 1) + ' FCFA/kg • ' + fr(p.formule.prixSac[i]) + ' FCFA le sac de 50 kg']),
        kv: [['Coût d’une poulette à 18 semaines', fr(p.coutPoulette) + ' FCFA'], ['Coût de production d’un œuf', fr(p.coutOeuf, 1) + ' FCFA'], ['Coût d’une plaquette de 30 œufs', fr(p.coutPlaq) + ' FCFA'],
          ['Prix de vente retenu', fr(p.prixPlaq) + ' FCFA/plaquette'], ['Marge brute par plaquette', fr(p.margePlaq) + ' FCFA'], ['Marge brute du cycle', eu(p.margePlaq * p.plaquettes)],
          ['Chiffre d’affaires annuel', eu(p.caAn)], ['Marge annuelle après amortissement (estimation)', eu(p.margeAn)],
          ['Prémix MG2MIX nécessaire', fr(p.premix.chair) + ' kg chair-poulette + ' + fr(p.premix.ponte) + ' kg ponte par cycle']] });
    }
    const titre = o.valide ? 'VOTRE SIMULATION VALIDÉE AVEC SOLEVO' : 'VOS COÛTS AVEC LES PRÉMIX SOLEVO MG2MIX';
    const intro = (o.nom ? 'Bonjour ' + o.nom + ',' : 'Bonjour,') + ' ' + (o.valide ? 'voici la simulation commentée et validée avec votre conseiller SOLEVO.' : 'voici vos coûts calculés avec les prix que vous avez indiqués' + (o.localite ? ' (' + o.localite + ')' : '') + '. Votre conseiller SOLEVO les commentera avec vous lors de l’entretien.');
    const note = 'Estimation indicative issue du modèle SOLEVO, avant impôts et frais financiers ; les prix appliqués lors de l’achat peuvent varier. 1 € = 656 FCFA.';
    const contact = 'SOLEVO Côte d’Ivoire – Elvis LEHITO, Key Account Manager FEED – WhatsApp +225 07 10 47 14 02 – elvis.lehito@solevogroup.com';
    const text = ['*' + titre + '*', intro].concat(sec.flatMap((x) => ['', '*' + x.t + '*', '_' + x.f + '_'].concat(x.ph.map(([a, b]) => '• ' + a + ' : ' + b), x.kv.map(([a, b]) => '• ' + a + ' : *' + b + '*')))).concat(['', note, contact]).join('\n');
    const H = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const html = `<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;max-width:640px"><div style="background:#002060;color:#fff;padding:14px 16px;border-bottom:4px solid #37A028"><b style="font-size:17px">${H(titre)}</b><br><span style="font-size:13px">Stand SOLEVO – SELAB SHOW 2026 • Programme SISSÊ~KABA</span></div><p>${H(intro)}</p>` +
      sec.map((x) => `<h3 style="color:#002060;margin:16px 0 4px">${H(x.t)}</h3><p style="margin:0 0 6px;color:#37A028"><b>${H(x.f)}</b></p><table style="border-collapse:collapse;width:100%;font-size:13px">${x.ph.map(([a, b]) => `<tr><td style="padding:3px 6px;border-bottom:1px solid #e3e8ef">${H(a)}</td><td style="padding:3px 6px;border-bottom:1px solid #e3e8ef;text-align:right">${H(b)}</td></tr>`).join('')}</table><table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:8px">${x.kv.map(([a, b]) => `<tr><td style="padding:4px 6px;border-bottom:1px solid #e3e8ef">${H(a)}</td><td style="padding:4px 6px;border-bottom:1px solid #e3e8ef;text-align:right"><b>${H(b)}</b></td></tr>`).join('')}</table>`).join('') +
      `<p style="font-size:12px;color:#666;margin-top:14px"><i>${H(note)}</i></p><p style="color:#002060"><b>${H(contact)}</b></p></div>`;
    return { text, html, titre };
  }
  /* saisie à partir de la veille d'un visiteur : veille = { mais, soja, son, coq, huile (FCFA/l), pxChair, pxChairKg, pxPonte, pxPonteKg, plaq, poulet } */
  function inputFromVeille(vx, eff) {
    vx = vx || {}; eff = eff || {}; const n = (k) => +vx[k] || 0; const prix = {};
    ['mais', 'soja', 'son', 'coq'].forEach((k) => { if (n(k)) prix[k] = n(k); });
    if (n('huile')) prix.huile = Math.round(n('huile') / 0.91);
    [['pxChair', 'pxChairKg'], ['pxPonte', 'pxPonteKg']].forEach(([k, kk]) => { const kg = n(kk); if (n(k)) prix[k] = Math.round(kg > 0 ? n(k) / kg : n(k) > 5000 ? n(k) / 25 : n(k)); });
    const ch = {}, po = {};
    if (n('poulet')) ch.prixVente = n('poulet'); if (n('plaq')) po.prixPlaq = n('plaq');
    if (+eff.chair > 0) ch.effectif = Math.round(+eff.chair); if (+eff.pondeuse > 0) po.effectif = Math.round(+eff.pondeuse);
    const sps = []; if (+eff.chair > 0 || n('poulet')) sps.push('chair'); if (+eff.pondeuse > 0 || n('plaq')) sps.push('ponte');
    return { inp: { prix, chair: Object.assign({ formule: 1 }, ch), ponte: Object.assign({ formule: 1 }, po), batiment: 'neuf' }, sps: sps.length ? sps : ['chair', 'ponte'], veille: Object.keys(prix) };
  }

  /* sensibilité : marge unitaire selon le prix du maïs (lignes) et le prix de vente (colonnes) */
  function sens(inp, sp, dm, dv) {
    dm = dm || [-0.2, -0.1, 0, 0.1, 0.2]; dv = dv || [-0.1, -0.05, 0, 0.05, 0.1];
    const cId = +((inp.chair || {}).formule) || 1, pId = +((inp.ponte || {}).formule) || 1;
    const mais0 = +((inp.prix || {}).mais) > 0 ? +inp.prix.mais : BASE.prix.mais;
    const rows = dm.map((m) => {
      const v = evalAll(Object.assign({}, inp, { prix: Object.assign({}, inp.prix, { mais: mais0 * (1 + m) }) }), cId, pId);
      const g = (s, r) => N(v[s + '!' + r]);
      const cout = sp === 'ponte' ? g(CO, 'C58') : g(CPC, 'C41'), prix = sp === 'ponte' ? g(CO, 'C59') : g(CPC, 'C44');
      return { m, mais: mais0 * (1 + m), cout, prix, marges: dv.map((w) => prix * (1 + w) - cout) };
    });
    return { dm, dv, rows, prix0: rows[0] ? rows[0].prix : 0 };
  }
  root.SKF = { sens, load, loaded, compute, exportXlsx, report, inputFromVeille, MP, PAR, CHAIR_F, PONTE_F, EUR, base: () => BASE };
})(typeof window !== 'undefined' ? window : globalThis);
