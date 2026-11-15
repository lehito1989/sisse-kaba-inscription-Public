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
    o = o || {}; const sps = o.sps && (o.sps.length || o.porc) ? o.sps : ['chair', 'ponte'];
    const sec = [];
    if (sps.includes('chair')) {
      const c = R.chair;
      sec.push({ t: '🐔 POULET DE CHAIR – ' + fr(c.effectif) + ' sujets, bande de ' + fr(c.duree) + ' jours', f: 'Aliment formulé ' + c.formule.nom + ' + MG2MIX 2 %',
        ph: c.phases.map((p, i) => [p.nom + ' (' + p.age + ')', fr(c.formule.prixKg[i], 1) + ' FCFA/kg • ' + fr(c.formule.prixSac[i]) + ' FCFA le sac de 50 kg']),
        kv: [['Coût de revient d’un poulet', fr(c.coutPoulet) + ' FCFA (' + fr(c.coutKgVif) + ' FCFA/kg vif)'], ['Prix de vente retenu', fr(c.prixVente) + ' FCFA/poulet'],
          ['Marge brute par poulet', fr(c.margePoulet) + ' FCFA (' + fr(c.margePct * 100) + ' %)'], ['Marge brute de la bande', eu(c.margeBande)],
          ['Chiffre d’affaires annuel (' + fr(c.bandes) + ' bandes)', eu(c.caAn)], ['Marge annuelle après amortissement (estimation)', eu(c.margeAn)],
          ['Prémix MG2MIX chair nécessaire', fr(c.premix.chair) + ' kg par bande (≈ ' + fr(Math.ceil(c.premix.chair / 20)) + ' sacs de 20 kg)']] });
    }
    if (sps.includes('ponte')) {
      const p = R.ponte;
      sec.push({ t: '🥚 PONDEUSES – ' + fr(p.effectif) + ' poussins, cycle de 78 semaines', f: 'Aliment formulé ' + p.formule.nom + ' + MG2MIX 2 %',
        ph: p.phases.map((x, i) => [x.nom + ' (' + x.age + ')', fr(p.formule.prixKg[i], 1) + ' FCFA/kg • ' + fr(p.formule.prixSac[i]) + ' FCFA le sac de 50 kg']),
        kv: [['Coût d’une poulette à 18 semaines', fr(p.coutPoulette) + ' FCFA'], ['Coût de production d’un œuf', fr(p.coutOeuf, 1) + ' FCFA'], ['Coût d’une plaquette de 30 œufs', fr(p.coutPlaq) + ' FCFA'],
          ['Prix de vente retenu', fr(p.prixPlaq) + ' FCFA/plaquette'], ['Marge brute par plaquette', fr(p.margePlaq) + ' FCFA'], ['Marge brute du cycle', eu(p.margePlaq * p.plaquettes)],
          ['Chiffre d’affaires annuel', eu(p.caAn)], ['Marge annuelle après amortissement (estimation)', eu(p.margeAn)],
          ['Prémix MG2MIX nécessaire', fr(p.premix.chair) + ' kg chair-poulette + ' + fr(p.premix.ponte) + ' kg ponte par cycle (≈ ' + fr(Math.ceil(p.premix.chair / 20)) + ' + ' + fr(Math.ceil(p.premix.ponte / 20)) + ' sacs de 20 kg)']] });
    }
    if (o.porc) {
      const L = porcCompute(R.prix), n = +o.porcN || 0;
      sec.push({ t: '🐖 PORCS' + (n ? ' – ' + fr(n) + ' porcs par an' : ''), f: 'Formules du livret MG2MIX porc + CMV Porc MG2MIX 2 % (1 sac de 20 kg = 1 tonne)',
        ph: L.map((f) => [f.nom, f.prixKg != null ? fr(f.prixKg, 1) + ' FCFA/kg • ' + fr(f.prixSac) + ' FCFA le sac de 50 kg' : (f.horsPremix != null ? fr(f.horsPremix, 1) + ' FCFA/kg hors prémix' : 'à chiffrer avec votre conseiller')]),
        kv: (() => { const P = porcProd(R.prix, o.porcVif ? { prixVif: o.porcVif } : {}), K = [];
          if (P.cout != null) K.push(['Coût de revient d’un porc de ' + fr(P.q.poidsS) + ' kg (repères à confirmer)', fr(P.cout) + ' FCFA (' + fr(P.coutKgVif) + ' FCFA/kg vif)'], ['Prix de vente au seuil de rentabilité', fr(P.seuil) + ' FCFA/kg vif'], ['Marge par porc à ' + fr(P.q.prixVif) + ' FCFA/kg vif', eu(P.marge)]);
          if (n) K.push(['CMV Porc MG2MIX nécessaire', fr(Math.ceil(n * P.kgAl / 1000)) + ' sacs de 20 kg par an']);
          return K; })() });
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
    [['pxChair', 'pxChairKg'], ['pxPonte', 'pxPonteKg'], ['pxPorc', 'pxPorcKg']].forEach(([k, kk]) => { const kg = n(kk); if (n(k)) prix[k] = Math.round(kg > 0 ? n(k) / kg : n(k) > 5000 ? n(k) / 20 : n(k)); });
    const ch = {}, po = {};
    if (n('poulet')) ch.prixVente = n('poulet'); if (n('plaq')) po.prixPlaq = n('plaq');
    if (+eff.chair > 0) ch.effectif = Math.round(+eff.chair); if (+eff.pondeuse > 0) po.effectif = Math.round(+eff.pondeuse);
    const sps = []; if (+eff.chair > 0 || n('poulet')) sps.push('chair'); if (+eff.pondeuse > 0 || n('plaq')) sps.push('ponte');
    const porc = +eff.porc > 0 || n('pxPorc') > 0 || n('porcVif') > 0;
    return { inp: { prix, chair: Object.assign({ formule: 1 }, ch), ponte: Object.assign({ formule: 1 }, po), batiment: 'neuf' }, sps: sps.length ? sps : (porc ? [] : ['chair', 'ponte']), porc, porcN: +eff.porc || 0, porcVif: n('porcVif') || null, veille: Object.keys(prix) };
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

  /* Livret de formules MG2MIX porc (SOLEVO Côte d'Ivoire, 13/01/2026) : CMV Porc MG2MIX 2 %, 20 kg par tonne à tous les stades */
  const PORC_MP = [['mais', 'Maïs'], ['soja', 'Tourteau de soja'], ['son', 'Son de blé'], ['coq', 'Coquillages'], ['palmiste', 'Tourteau de palmiste'], ['fibrePalm', 'Fibre de pression de palmiste'], ['pxPorc', 'CMV Porc MG2MIX 2 %']];
  const PORC_F = [
    { id: 'porcelet', nom: 'Porcelets', kg: { mais: 625, soja: 200, son: 150, coq: 5, pxPorc: 20 } },
    { id: 'charcutier', nom: 'Charcutiers', kg: { mais: 606, soja: 220, son: 145, coq: 9, pxPorc: 20 } },
    { id: 'charcPalm', nom: 'Charcutiers (avec palmiste)', kg: { mais: 620, son: 180, soja: 147, fibrePalm: 25, coq: 8, pxPorc: 20 } },
    { id: 'gestante', nom: 'Truie gestante', kg: { mais: 566, son: 250, soja: 76, palmiste: 72, coq: 16, pxPorc: 20 } },
    { id: 'allaitante', nom: 'Truie allaitante', kg: { mais: 606, soja: 220, son: 145, coq: 9, pxPorc: 20 } }];
  function porcCompute(prix) {
    const P = Object.assign({}, BASE ? BASE.prix : {}, prix || {});
    return PORC_F.map((f) => {
      const rows = PORC_MP.filter(([k]) => f.kg[k]).map(([k, nom]) => ({ k, nom, kg: f.kg[k], prix: +P[k] > 0 ? +P[k] : null }));
      const missing = rows.filter((r) => r.prix == null).map((r) => r.nom);
      const prixKg = missing.length ? null : rows.reduce((a, r) => a + r.kg * r.prix, 0) / 1000;
      const horsPremix = rows.filter((r) => r.k !== 'pxPorc').every((r) => r.prix != null) ? rows.filter((r) => r.k !== 'pxPorc').reduce((a, r) => a + r.kg * r.prix, 0) / 1000 : null;
      return { id: f.id, nom: f.nom, rows, missing, prixKg, prixSac: prixKg == null ? null : prixKg * 50, horsPremix };
    });
  }
  /* Registre des paramètres porc (relevé du 15/11/2026) : valeur, unité, date, source, statut.
     statut : 'publié' | 'annonce commerciale' | 'hypothèse à confirmer' | 'paramètre du simulateur' */
  const PORC_PARAMS = {
    effectif: { l: 'Porcs engraissés par bande', u: 'porcs', v: 100, date: '—', src: 'Saisie de l’éleveur', st: 'hypothèse à confirmer' },
    poidsE: { l: 'Poids d’entrée (porcelet sevré)', u: 'kg', v: 7, date: '2013', src: 'Étude EISMV, ferme AGP d’Azaguié : 7,14 kg au sevrage', st: 'publié' },
    ageE: { l: 'Âge d’entrée', u: 'jours', v: 30, date: '2013', src: 'Étude EISMV, ferme AGP d’Azaguié : 30,6 jours', st: 'publié' },
    poidsT: { l: 'Poids de passage porcelet → charcutier', u: 'kg', v: 25, date: '—', src: 'Repère de conduite', st: 'hypothèse à confirmer' },
    poidsS: { l: 'Poids de vente', u: 'kg vif', v: 90, date: '15/11/2026', src: 'Récapitulatif de recherche du 15/11/2026', st: 'hypothèse à confirmer' },
    ageS: { l: 'Âge de vente', u: 'jours', v: 210, date: '15/11/2026', src: '7 mois, à vérifier selon le poids d’entrée et la croissance', st: 'hypothèse à confirmer' },
    icP: { l: 'Indice de consommation porcelet', u: 'kg aliment / kg gain', v: 2.0, date: '—', src: 'Repère zootechnique', st: 'hypothèse à confirmer' },
    icC: { l: 'Indice de consommation charcutier', u: 'kg aliment / kg gain', v: 2.85, date: '—', src: 'Déduit du paramètre « 250 kg d’aliment par porc » du simulateur', st: 'paramètre du simulateur' },
    mort: { l: 'Mortalité post-sevrage → vente', u: '%', v: 5, date: '15/11/2026', src: 'Hypothèse de simulation (les pertes PPA 2024 sont exceptionnelles)', st: 'hypothèse à confirmer' },
    prixPorcelet: { l: 'Achat du porcelet', u: 'FCFA / tête', v: 30000, date: '22/05/2025', src: 'Annonce EPAP N’gattakro, Yamoussoukro (poids et âge non précisés) ; 25 000 FCFA à La Mé le 21/02/2026', st: 'annonce commerciale' },
    prixVif: { l: 'Prix de vente', u: 'FCFA / kg vif', v: 1500, date: 'non daté', src: 'Annonces commerciales non confirmées ; scénarios 1 300 / 1 500 / 1 700', st: 'hypothèse à confirmer' },
    autres: { l: 'Autres charges par porc mis en place', u: 'FCFA', v: 0, date: '—', src: 'Soins, eau, énergie, main-d’œuvre : à relever chez l’éleveur', st: 'hypothèse à confirmer' },
    truies: { l: 'Truies en production', u: 'truies', v: 10, date: '—', src: 'Saisie de l’éleveur', st: 'hypothèse à confirmer' },
    portees: { l: 'Portées par truie et par an', u: 'portées', v: 1.8, date: '—', src: 'Pas de donnée ivoirienne (1,56 au Bénin dans l’étude EISMV)', st: 'hypothèse à confirmer' },
    sevres: { l: 'Porcelets sevrés par portée', u: 'porcelets', v: 8.37, date: '2013', src: 'Étude EISMV, ferme AGP d’Azaguié', st: 'publié' },
    gestJ: { l: 'Durée de gestation', u: 'jours', v: 114, date: '—', src: 'Repère biologique de l’espèce', st: 'publié' },
    gestKg: { l: 'Aliment truie gestante', u: 'kg / jour', v: 2.5, date: '—', src: 'Repère zootechnique', st: 'hypothèse à confirmer' },
    lactJ: { l: 'Durée de lactation', u: 'jours', v: 30, date: '2013', src: 'Sevrage à 30,6 jours (étude EISMV)', st: 'publié' },
    lactKg: { l: 'Aliment truie allaitante', u: 'kg / jour', v: 5, date: '—', src: 'Repère zootechnique', st: 'hypothèse à confirmer' },
    videJ: { l: 'Intervalle sevrage → saillie', u: 'jours', v: 7, date: '—', src: 'Repère zootechnique (aliment gestante)', st: 'hypothèse à confirmer' },
    autresTruie: { l: 'Autres charges par truie et par an', u: 'FCFA', v: 0, date: '—', src: 'Soins, verrat ou IA, logement : à relever', st: 'hypothèse à confirmer' }
  };
  const PORC_PRIX_SRC = {
    mais: { date: 'registre SOLEVO', src: 'Registre terrain du classeur de formulation SOLEVO', st: 'paramètre du simulateur' },
    soja: { date: 'registre SOLEVO', src: 'Registre terrain du classeur de formulation SOLEVO', st: 'paramètre du simulateur' },
    son: { date: 'registre SOLEVO', src: 'Registre terrain du classeur de formulation SOLEVO', st: 'paramètre du simulateur' },
    coq: { date: 'registre SOLEVO', src: 'Registre terrain du classeur de formulation SOLEVO', st: 'paramètre du simulateur' },
    pxPorc: { date: 'simulateur', src: 'Prix du sac de prémix porcin du simulateur ÷ 20 kg', st: 'paramètre du simulateur' },
    palmiste: { date: 'à confirmer', src: 'Annonces : 45 FCFA/kg départ San Pedro, 65 FCFA/kg livré Abidjan, 100 FCFA/kg FOB ; prix rendu ferme à confirmer', st: 'hypothèse à confirmer' },
    fibrePalm: { date: 'à confirmer', src: 'Aucun prix fiable ; nature exacte du produit à vérifier auprès du fournisseur', st: 'hypothèse à confirmer' }
  };
  const PALM_ANNONCES = [{ v: 45, l: '45 FCFA/kg – départ magasin San Pedro' }, { v: 65, l: '65 FCFA/kg – livré Abidjan' }, { v: 100, l: '100 FCFA/kg – offre FOB Côte d’Ivoire' }];
  const pdef = (k) => PORC_PARAMS[k].v;
  /* engraissement : coût par porc vendu, en comptant les porcs morts (porcelet + charges) et l’aliment qu’ils ont mangé (mort en moyenne à mi-parcours) */
  function porcProd(prix, q) {
    const Q = {}; Object.keys(PORC_PARAMS).forEach((k) => { Q[k] = q && q[k] !== '' && q[k] != null ? +q[k] : pdef(k); }); Q.formC = (q && q.formC) || 'charcutier';
    const L = porcCompute(prix), fP = L.find((f) => f.id === 'porcelet'), fC = L.find((f) => f.id === Q.formC) || L.find((f) => f.id === 'charcutier');
    const m = Math.min(0.9, Math.max(0, Q.mort / 100)), alP = Math.max(0, (Q.poidsT - Q.poidsE) * Q.icP), alC = Math.max(0, (Q.poidsS - Q.poidsT) * Q.icC), kgAl = alP + alC, gain = Q.poidsS - Q.poidsE;
    const coutAl = fP.prixKg != null && fC.prixKg != null ? alP * fP.prixKg + alC * fC.prixKg : null;
    const N = Q.effectif, vendus = N * (1 - m);
    const totalBande = coutAl != null ? N * (Q.prixPorcelet + Q.autres) + vendus * coutAl + N * m * 0.5 * coutAl : null;
    const cout = totalBande != null && vendus > 0 ? totalBande / vendus : null;
    const coutMorts = coutAl != null && vendus > 0 ? N * m * (Q.prixPorcelet + Q.autres + 0.5 * coutAl) / vendus : null;
    const ca = Q.prixVif * Q.poidsS, marge = cout != null ? ca - cout : null;
    const scen = [1300, 1500, 1700].map((px) => ({ px, ca: px * Q.poidsS, marge: cout != null ? px * Q.poidsS - cout : null, margeBande: cout != null ? (px * Q.poidsS - cout) * vendus : null }));
    return { q: Q, fP, fC, alP, alC, kgAl, ic: gain > 0 ? kgAl / gain : null, gmq: Q.ageS > Q.ageE ? gain / (Q.ageS - Q.ageE) * 1000 : null, duree: Q.ageS - Q.ageE,
      coutAl, cout, coutMorts, coutKgVif: cout != null ? cout / Q.poidsS : null, seuil: cout != null ? cout / Q.poidsS : null, partAl: cout ? (vendus * coutAl + N * m * 0.5 * coutAl) / totalBande : null,
      vendus: Math.round(vendus), ca, marge, margePct: marge != null && ca ? marge / ca : null, margeBande: marge != null ? marge * vendus : null, scen,
      premixKg: (vendus + N * m * 0.5) * kgAl * 0.02, sacs: Math.ceil((vendus + N * m * 0.5) * kgAl / 1000) };
  }
  /* naisseur : porcelets sevrés par truie et par an, coût d’un porcelet sevré (aliments gestante + allaitante) */
  function porcNaisseur(prix, q) {
    const Q = {}; ['truies', 'portees', 'sevres', 'gestJ', 'gestKg', 'lactJ', 'lactKg', 'videJ', 'autresTruie'].forEach((k) => { Q[k] = q && q[k] !== '' && q[k] != null ? +q[k] : pdef(k); });
    const L = porcCompute(prix), fG = L.find((f) => f.id === 'gestante'), fA = L.find((f) => f.id === 'allaitante');
    const kgG = Q.portees * (Q.gestJ + Q.videJ) * Q.gestKg, kgA = Q.portees * Q.lactJ * Q.lactKg, sevresAn = Q.portees * Q.sevres;
    const coutG = fG.prixKg != null ? kgG * fG.prixKg : null, coutA = fA.prixKg != null ? kgA * fA.prixKg : null;
    const coutTruie = coutG != null && coutA != null ? coutG + coutA + Q.autresTruie : null;
    return { q: Q, fG, fA, kgG, kgA, kgAn: kgG + kgA, sevresAn, sevresTot: sevresAn * Q.truies, cycleJ: Q.gestJ + Q.lactJ + Q.videJ, porteesMax: 365 / (Q.gestJ + Q.lactJ + Q.videJ),
      coutG, coutA, coutTruie, coutPorcelet: coutTruie != null && sevresAn > 0 ? coutTruie / sevresAn : null, sacsAn: Math.ceil((kgG + kgA) * Q.truies / 1000) };
  }
  root.SKF = { porcNaisseur, PORC_PARAMS, PORC_PRIX_SRC, PALM_ANNONCES, porcProd, PORC_F, PORC_MP, porcCompute, sens, load, loaded, compute, exportXlsx, report, inputFromVeille, MP, PAR, CHAIR_F, PONTE_F, EUR, base: () => BASE };
})(typeof window !== 'undefined' ? window : globalThis);
