/* SISSÊ~KABA — moteur de calcul : lit le modèle Excel, évalue ses formules, réécrit les valeurs.
   Fonctionne dans le navigateur et sous Node (tests). */
(function (root) {
  'use strict';

  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function decode(s) {
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (m, e) => {
      if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return ENT[e] !== undefined ? ENT[e] : m;
    });
  }
  function encode(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* ---------- lecture des feuilles ---------- */
  const CELL_RE = /<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  /* décale les références relatives d’une formule partagée (dc colonnes, dr lignes) */
  function shiftFormula(f, dc, dr) {
    let out = '', i = 0;
    while (i < f.length) {
      if (f[i] === '"') { const j = f.indexOf('"', i + 1); const k = j < 0 ? f.length : j + 1; out += f.slice(i, k); i = k; continue; }
      let j = f.indexOf('"', i); if (j < 0) j = f.length;
      out += f.slice(i, j).replace(/((?:'(?:[^']|'')+'|[A-Za-zÀ-ÿ_][\wÀ-ÿ.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g, (m, sh, dcol, col, drow, row) => {
        const c = dcol ? col : colStr(colNum(col) + dc); const r = drow ? row : String(+row + dr);
        return (sh || '') + dcol + c + drow + r;
      });
      i = j;
    }
    return out;
  }
  function parseSheet(xml, shared) {
    const cells = {}, masters = {};
    let m;
    CELL_RE.lastIndex = 0;
    while ((m = CELL_RE.exec(xml))) {
      const ref = m[1] + m[2];
      const attrs = m[3] || '';
      const inner = m[4] || '';
      const t = (attrs.match(/\bt="([^"]*)"/) || [])[1] || 'n';
      const fm = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
      const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
      let value = null;
      if (t === 'inlineStr') {
        const ts = inner.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [];
        value = decode(ts.map((x) => x.replace(/<[^>]+>/g, '')).join(''));
      } else if (t === 's' && vm) value = shared[+vm[1]];
      else if (t === 'str' && vm) value = decode(vm[1]);
      else if (t === 'b' && vm) value = vm[1] === '1';
      else if (vm && vm[1] !== '') value = parseFloat(vm[1]);
      let f = fm ? decode(fm[1]) : null;
      const sh = inner.match(/<f\b([^>]*?)(\/>|>)/);
      if (sh && /t="shared"/.test(sh[1])) {
        const si = (sh[1].match(/si="(\d+)"/) || [])[1];
        if (f) masters[si] = { f, c: colNum(m[1]), r: +m[2] };
        else if (masters[si]) f = shiftFormula(masters[si].f, colNum(m[1]) - masters[si].c, +m[2] - masters[si].r);
      }
      cells[ref] = { f, value };
    }
    return cells;
  }

  async function loadWorkbook(zip) {
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rels = {};
    relXml.replace(/<Relationship ([^>]*)\/>/g, (m, a) => {
      const id = (a.match(/Id="([^"]+)"/) || [])[1];
      let tg = (a.match(/Target="([^"]+)"/) || [])[1];
      if (tg) tg = tg.replace(/^\//, '');
      if (tg && !tg.startsWith('xl/')) tg = 'xl/' + tg;
      rels[id] = tg;
      return m;
    });
    let shared = [];
    const ssf = zip.file('xl/sharedStrings.xml');
    if (ssf) {
      const ss = await ssf.async('string');
      shared = (ss.match(/<si>[\s\S]*?<\/si>/g) || []).map((si) =>
        decode((si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map((x) => x.replace(/<[^>]+>/g, '')).join('')));
    }
    const sheets = {};
    const order = [];
    const re = /<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*\/>/g;
    let m;
    while ((m = re.exec(wbXml))) {
      const name = decode(m[1]);
      const path = rels[m[2]];
      const xml = await zip.file(path).async('string');
      sheets[name] = { path, xml, cells: parseSheet(xml, shared) };
      order.push(name);
    }
    return { zip, sheets, order, shared };
  }

  /* ---------- formules ---------- */
  function tokenize(src) {
    const toks = [];
    let i = 0;
    while (i < src.length) {
      const c = src[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '{') {
        const j = src.indexOf('}', i); const body = src.slice(i + 1, j); const items = [];
        body.replace(/"((?:[^"]|"")*)"|(-?\d+(?:\.\d+)?)/g, (m, s1, n1) => { items.push(s1 !== undefined ? s1.replace(/""/g, '"') : parseFloat(n1)); return m; });
        toks.push({ t: 'arr', v: items }); i = j + 1; continue;
      }
      if (c === '"') {
        let j = i + 1, s = '';
        while (j < src.length) {
          if (src[j] === '"' && src[j + 1] === '"') { s += '"'; j += 2; continue; }
          if (src[j] === '"') break;
          s += src[j++];
        }
        toks.push({ t: 'str', v: s }); i = j + 1; continue;
      }
      const ref = src.slice(i).match(/^(?:('(?:[^']|'')+'|[A-Za-zÀ-ÿ_][\wÀ-ÿ.]*)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(?:(?:'(?:[^']|'')+'|[A-Za-zÀ-ÿ_][\wÀ-ÿ.]*)!)?(\$?[A-Z]{1,3}\$?\d+))?/);
      if (ref && !/^[A-Z]+\(/.test(src.slice(i))) {
        let sheet = ref[1] ? ref[1].replace(/^'|'$/g, '').replace(/''/g, "'") : null;
        toks.push({ t: 'ref', sheet, a: ref[2].replace(/\$/g, ''), b: ref[3] ? ref[3].replace(/\$/g, '') : null });
        i += ref[0].length; continue;
      }
      const num = src.slice(i).match(/^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/);
      if (num) { toks.push({ t: 'num', v: parseFloat(num[0]) }); i += num[0].length; continue; }
      const bl = src.slice(i).match(/^(TRUE|FALSE)(?![\w(])/);
      if (bl) { toks.push({ t: 'bool', v: bl[1] === 'TRUE' }); i += bl[0].length; continue; }
      const fn = src.slice(i).match(/^([A-Z][A-Z0-9.]*)\(/);
      if (fn) { toks.push({ t: 'fn', v: fn[1] }); i += fn[0].length; continue; }
      const op = src.slice(i).match(/^(<=|>=|<>|[-+*/^&=<>(),%])/);
      if (op) { toks.push({ t: 'op', v: op[1] }); i += op[0].length; continue; }
      throw new Error('Formule non reconnue : ' + src + ' @' + i);
    }
    return toks;
  }

  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const isOp = (v) => toks[p] && toks[p].t === 'op' && toks[p].v === v;
    function expr() { return cmp(); }
    function cmp() {
      let l = concat();
      while (toks[p] && toks[p].t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(toks[p].v)) {
        const o = toks[p++].v; const r = concat(); l = { k: 'bin', o, l, r };
      }
      return l;
    }
    function concat() { let l = add(); while (isOp('&')) { p++; l = { k: 'bin', o: '&', l, r: add() }; } return l; }
    function add() { let l = mul(); while (isOp('+') || isOp('-')) { const o = toks[p++].v; l = { k: 'bin', o, l, r: mul() }; } return l; }
    function mul() { let l = pow(); while (isOp('*') || isOp('/')) { const o = toks[p++].v; l = { k: 'bin', o, l, r: pow() }; } return l; }
    function pow() { let l = unary(); while (isOp('^')) { p++; l = { k: 'bin', o: '^', l, r: unary() }; } return l; }
    function unary() {
      if (isOp('-')) { p++; return { k: 'neg', e: unary() }; }
      if (isOp('+')) { p++; return unary(); }
      let e = prim();
      if (isOp('%')) { p++; e = { k: 'bin', o: '/', l: e, r: { k: 'num', v: 100 } }; }
      return e;
    }
    function prim() {
      const t = peek();
      if (!t) throw new Error('Fin inattendue : ' + src);
      if (t.t === 'num') { p++; return { k: 'num', v: t.v }; }
      if (t.t === 'str') { p++; return { k: 'str', v: t.v }; }
      if (t.t === 'arr') { p++; return { k: 'arr', v: t.v }; }
      if (t.t === 'bool') { p++; return { k: 'bool', v: t.v }; }
      if (t.t === 'ref') { p++; return { k: 'ref', sheet: t.sheet, a: t.a, b: t.b }; }
      if (t.t === 'fn') {
        p++; const args = [];
        if (!isOp(')')) { args.push(expr()); while (isOp(',')) { p++; args.push(expr()); } }
        if (!isOp(')')) throw new Error('Parenthèse manquante : ' + src);
        p++; return { k: 'fn', n: t.v, args };
      }
      if (isOp('(')) { p++; const e = expr(); if (!isOp(')')) throw new Error(') attendue : ' + src); p++; return e; }
      throw new Error('Jeton inattendu ' + JSON.stringify(t) + ' dans ' + src);
    }
    const ast = expr();
    if (p !== toks.length) throw new Error('Formule incomplète : ' + src);
    return ast;
  }

  function colNum(c) { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function colStr(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function splitRef(r) { const m = r.match(/^([A-Z]+)(\d+)$/); return [colNum(m[1]), +m[2]]; }
  const XE = (c) => ({ __e: c || '#N/A' });
  const isErr = (v) => (v && typeof v === 'object' && !Array.isArray(v) && v.__e) || (typeof v === 'number' && !isFinite(v));
  function fmtText(v, fmt) {
    const x = typeof v === 'number' ? v : parseFloat(v); if (!isFinite(x)) return String(v ?? '');
    const dm = fmt.match(/0[.,](0+)/); const d = dm ? dm[1].length : 0; const grp = /#[ ,.\u00a0\u202f]##/.test(fmt);
    let s = Math.abs(x).toFixed(d); let [a, b] = s.split('.'); if (grp) a = a.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (x < 0 ? '-' : '') + a + (b ? ',' + b : '');
  }
  const num = (v) => (v === null || v === undefined || v === '' ? 0 : typeof v === 'boolean' ? (v ? 1 : 0) : typeof v === 'number' ? v : (isNaN(+v) ? 0 : +v));
  function eqv(a, b) {
    if (a === null || a === undefined) a = typeof b === 'string' ? '' : 0;
    if (b === null || b === undefined) b = typeof a === 'string' ? '' : 0;
    if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
    return a === b;
  }

  /* évalue le classeur ; overrides = { 'Feuille!A1': valeur } */
  function evaluate(wb, overrides) {
    overrides = overrides || {};
    const cache = {}, busy = {}, asts = {};
    function get(sheet, ref) {
      const key = sheet + '!' + ref;
      if (key in overrides) return overrides[key];
      if (key in cache) return cache[key];
      const sh = wb.sheets[sheet];
      if (!sh) throw new Error('Feuille inconnue ' + sheet);
      const c = sh.cells[ref];
      if (!c) return null;
      if (!c.f) return c.value;
      if (busy[key]) throw new Error('Référence circulaire ' + key);
      busy[key] = true;
      const ast = asts[key] || (asts[key] = parse(c.f));
      const v = ev(ast, sheet);
      busy[key] = false;
      cache[key] = v;
      return v;
    }
    function range(sheet, a, b) {
      const [c1, r1] = splitRef(a), [c2, r2] = splitRef(b), out = [];
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(get(sheet, colStr(c) + r));
      out.cols = Math.abs(c2 - c1) + 1;
      return out;
    }
    function ev(n, sheet) {
      switch (n.k) {
        case 'num': return n.v;
        case 'str': return n.v;
        case 'ref': return n.b ? range(n.sheet || sheet, n.a, n.b) : get(n.sheet || sheet, n.a);
        case 'bool': return n.v;
        case 'arr': { const a = n.v.slice(); a.cols = a.length; return a; }
        case 'neg': { const v = ev(n.e, sheet); return isErr(v) ? v : -num(v); }
        case 'bin': {
          const l = ev(n.l, sheet), r = ev(n.r, sheet);
          if (isErr(l)) return l; if (isErr(r)) return r;
          switch (n.o) {
            case '+': return num(l) + num(r);
            case '-': return num(l) - num(r);
            case '*': return num(l) * num(r);
            case '/': return num(r) === 0 ? NaN : num(l) / num(r);
            case '^': return Math.pow(num(l), num(r));
            case '&': return String(l ?? '') + String(r ?? '');
            case '=': return eqv(l, r);
            case '<>': return !eqv(l, r);
            case '<': return num(l) < num(r);
            case '>': return num(l) > num(r);
            case '<=': return num(l) <= num(r);
            case '>=': return num(l) >= num(r);
          }
          break;
        }
        case 'fn': {
          const A = n.args;
          const flat = () => A.map((a) => ev(a, sheet)).flat();
          switch (n.n) {
            case 'IF': { const c = ev(A[0], sheet); return (typeof c === 'string' ? c !== '' : num(c) !== 0 && c !== false) ? ev(A[1], sheet) : (A.length > 2 ? ev(A[2], sheet) : false); }
            case 'SUM': return flat().reduce((s, v) => s + (typeof v === 'number' ? v : 0), 0);
            case 'MAX': { const v = flat().filter((x) => typeof x === 'number'); return v.length ? Math.max(...v) : 0; }
            case 'MIN': { const v = flat().filter((x) => typeof x === 'number'); return v.length ? Math.min(...v) : 0; }
            case 'ABS': return Math.abs(num(ev(A[0], sheet)));
            case 'ROUNDUP': { const x = num(ev(A[0], sheet)), d = num(ev(A[1], sheet)), f = Math.pow(10, d); const y = Math.abs(x) * f; const rr = Math.ceil(y - 1e-9); return Math.sign(x) * rr / f; }
            case 'ROUND': { const x = num(ev(A[0], sheet)), f = Math.pow(10, num(ev(A[1], sheet))); return Math.round(x * f) / f; }
            case 'COUNTIF': { const vals = ev(A[0], sheet), crit = ev(A[1], sheet); return (Array.isArray(vals) ? vals : [vals]).filter((v) => eqv(v, crit)).length; }
            case 'AND': return flat().every((v) => num(v) !== 0 && v !== false);
            case 'OR': return flat().some((v) => num(v) !== 0 && v === true);
            case 'IFERROR': { let v; try { v = ev(A[0], sheet); } catch (e) { v = XE(); } return isErr(v) ? ev(A[1], sheet) : v; }
            case 'SUMPRODUCT': {
              const arrs = A.map((a) => { const v = ev(a, sheet); return Array.isArray(v) ? v : [v]; });
              const n0 = arrs[0].length; let s = 0;
              for (let i = 0; i < n0; i++) { let p = 1; for (const a of arrs) { const x = a[i]; p *= typeof x === 'number' ? x : (typeof x === 'boolean' ? (x ? 1 : 0) : 0); } s += p; }
              return s;
            }
            case 'VLOOKUP': {
              const key = ev(A[0], sheet), tb = ev(A[1], sheet), ci = num(ev(A[2], sheet)), approx = A.length > 3 ? num(ev(A[3], sheet)) !== 0 && ev(A[3], sheet) !== false : true;
              const w = tb.cols || 1, rows = tb.length / w;
              if (!approx) { for (let r = 0; r < rows; r++) if (eqv(tb[r * w], key)) return tb[r * w + ci - 1] ?? 0; return XE(); }
              let hit = -1; for (let r = 0; r < rows; r++) { const v = tb[r * w]; if (v === null || v === undefined) continue; if (num(v) <= num(key)) hit = r; else break; }
              return hit < 0 ? XE() : (tb[hit * w + ci - 1] ?? 0);
            }
            case 'MATCH': {
              const key = ev(A[0], sheet), arr = ev(A[1], sheet), ty = A.length > 2 ? num(ev(A[2], sheet)) : 1;
              const L = Array.isArray(arr) ? arr : [arr];
              if (ty === 0) { const i = L.findIndex((v) => eqv(v, key)); return i < 0 ? XE() : i + 1; }
              let hit = -1; L.forEach((v, i) => { if (v !== null && num(v) <= num(key)) hit = i; }); return hit < 0 ? XE() : hit + 1;
            }
            case 'INDEX': {
              const arr = ev(A[0], sheet), r = num(ev(A[1], sheet)), c = A.length > 2 ? num(ev(A[2], sheet)) : 1;
              const w = (arr && arr.cols) || 1; if (!Array.isArray(arr)) return arr;
              const idx = w === 1 ? r - 1 : (arr.length / w === 1 ? (c > 1 ? c : r) - 1 : (r - 1) * w + (c - 1));
              return arr[idx] ?? XE('#REF!');
            }
            case 'CHOOSE': { const i = num(ev(A[0], sheet)); if (isErr(ev(A[0], sheet)) || i < 1 || i >= A.length) return XE('#VALUE!'); return ev(A[i], sheet); }
            case 'TEXT': return fmtText(ev(A[0], sheet), String(ev(A[1], sheet)));
            case 'TRIM': return String(ev(A[0], sheet) ?? '').replace(/\s+/g, ' ').trim();
            case 'TODAY': { const d = new Date(); return Math.floor((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000); }
            default: throw new Error('Fonction non gérée : ' + n.n);
          }
        }
      }
      throw new Error('Nœud inconnu');
    }
    const all = {};
    for (const name of wb.order) {
      for (const ref of Object.keys(wb.sheets[name].cells)) all[name + '!' + ref] = get(name, ref);
    }
    for (const k in overrides) all[k] = overrides[k];
    return all;
  }

  /* réécrit valeurs saisies + résultats des formules dans le zip */
  function writeBack(wb, values, overrides) {
    overrides = overrides || {};
    for (const name of wb.order) {
      const sh = wb.sheets[name];
      const xml = sh.xml.replace(CELL_RE, (whole, col, row, attrs, inner) => {
        const ref = col + row, key = name + '!' + ref;
        const c = sh.cells[ref];
        const base = attrs.replace(/\s*\bt="[^"]*"/, '');
        if (key in overrides) {
          const v = overrides[key];
          if (typeof v === 'string') return `<c r="${ref}"${base} t="inlineStr"><is><t xml:space="preserve">${encode(v)}</t></is></c>`;
          return `<c r="${ref}"${base} t="n"><v>${v}</v></c>`;
        }
        if (c && c.f) {
          let v = values[key];
          const f = `<f>${encode(c.f)}</f>`;
          if (typeof v === 'string') return `<c r="${ref}"${base} t="str">${f}<v>${encode(v)}</v></c>`;
          if (typeof v === 'boolean') return `<c r="${ref}"${base} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
          if (v && v.__e) return `<c r="${ref}"${base} t="e">${f}<v>${v.__e}</v></c>`;
          if (typeof v !== 'number' || !isFinite(v)) return `<c r="${ref}"${base} t="e">${f}<v>#DIV/0!</v></c>`;
          return `<c r="${ref}"${base}>${f}<v>${+v.toPrecision(15)}</v></c>`;
        }
        return whole;
      });
      wb.zip.file(sh.path, xml);
    }
    return wb.zip;
  }

  root.SKX = { decode, encode, loadWorkbook, evaluate, writeBack, parse, parseSheet, shiftFormula, isErr };
})(typeof window !== 'undefined' ? window : globalThis);
