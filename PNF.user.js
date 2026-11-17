// ==UserScript==
// @name         PNF · Respuesta a oficios de Fiscalización
// @namespace    pnf.seremitt
// @version      0.3.25
// @description  Responde los oficios de Fiscalización desde el SGD: lee el oficio, arma el análisis técnico y el memo, y lo envía. Ventana, configuración, respaldo, usuarios, alertas, lectura de oficios, análisis técnico, memo, Excel y aportes de supervisores.
// @match        https://sgd.mtt.cl/*
// @noframes
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_xmlhttpRequest
// @grant        GM_download
// @grant        unsafeWindow
// @connect      generativelanguage.googleapis.com
// @connect      raw.githubusercontent.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @updateURL    https://raw.githubusercontent.com/paralepepipedo/PNF/main/PNF.user.js
// @downloadURL  https://raw.githubusercontent.com/paralepepipedo/PNF/main/PNF.user.js
// ==/UserScript==

(function () {
  'use strict';
  const VERSION = '0.3.25';
  const URL_ACTUALIZAR = 'https://raw.githubusercontent.com/paralepepipedo/PNF/main/PNF.user.js'; // debe coincidir con @updateURL
  const W = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
  if (window.top !== window.self || document.getElementById('pnf-host')) return;

  /* ═══════════════ 1 · Utilidades y almacenamiento del script (Tampermonkey) ═══════════════ */
  /* El SGD usa Prototype.js, que cambia JSON.stringify: convierte cada arreglo en TEXTO («"[{...}]"»). Por eso se desactiva un momento ese cambio al convertir. */
  const jsonStr = o => {
    const guardadas = []; [Array, Object, String, Number, Boolean].forEach(c => { const d = Object.getOwnPropertyDescriptor(c.prototype, 'toJSON'); if (d && d.configurable) { guardadas.push([c, d]); delete c.prototype.toJSON; } });
    try { return JSON.stringify(o); } finally { guardadas.forEach(g => { Object.defineProperty(g[0].prototype, 'toJSON', g[1]); }); }
  };
  // Datos que ya se guardaron con el defecto anterior (arreglos convertidos en texto): se recuperan al leerlos.
  function arr(v) { if (typeof v === 'string' && /^\s*\[/.test(v)) { try { const x = JSON.parse(v); if (Array.isArray(x)) return x; } catch (e) { /* no era un arreglo */ } } return v; }
  function normD(d) {
    d.catalogo = arr(d.catalogo); (Array.isArray(d.catalogo) ? d.catalogo : []).forEach(c => { c.kw = arr(c.kw); });
    if (d.cfg) d.cfg.personas = arr(d.cfg.personas);
    Object.keys(d.oficios || {}).forEach(e => { const o = d.oficios[e]; o.bloques = arr(o.bloques); (Array.isArray(o.bloques) ? o.bloques : []).forEach(b => { b.items = arr(b.items); }); });
    return d;
  }
  function normR(r) { r.admins = arr(r.admins); r.hist = arr(r.hist); return r; }
  /* Prototype.js agrega métodos ENUMERABLES a Array.prototype y pdf.js se niega a trabajar así: mientras corre una librería se vuelven no enumerables y después se dejan como estaban. */
  // texto de un error con el lugar del script donde ocurrió (para poder ubicarlo)
  function errTxt(x) { const m = (x && x.message) || String(x), st = String((x && x.stack) || '').split('\n').slice(1).map(l => l.trim().replace(/^at\s+/, '').replace(/\(?(?:blob:|chrome-extension:|https?:)[^)\s]*[\/]([^\/)\s]*:\d+:\d+)\)?/, '($1)')).filter(Boolean).slice(0, 3); return m + (st.length ? ' · en ' + st.join(' ← ') : ''); }
  /* Prototype.js 1.6 (el del SGD) además REEMPLAZA métodos nativos (map, filter, reduce, bind, Object.keys, Array.from…) por versiones antiguas que no hacen lo mismo
     (p. ej. Array.from no entiende Set/Map) y eso dejaba el memo PDF en blanco. Mientras corre una librería se ponen versiones estándar y después se devuelven las de la página. */
  const esNativa = f => typeof f === 'function' && /\[native code\]/.test(Function.prototype.toString.call(f));
  const estandar = () => [
    [Array.prototype, {
      find(f, t) { for (let i = 0; i < this.length; i++) if (f.call(t, this[i], i, this)) return this[i]; return undefined; },
      reverse() { for (let i = 0, j = this.length - 1; i < j; i++, j--) { const x = this[i]; this[i] = this[j]; this[j] = x; } return this; },
      entries() { const a = this; return (function* () { for (let i = 0; i < a.length; i++) yield [i, a[i]]; })(); },
      filter(f, t) { const r = []; for (let i = 0; i < this.length; i++) if (i in this && f.call(t, this[i], i, this)) r.push(this[i]); return r; },
      map(f, t) { const r = new Array(this.length); for (let i = 0; i < this.length; i++) if (i in this) r[i] = f.call(t, this[i], i, this); return r; },
      every(f, t) { for (let i = 0; i < this.length; i++) if (i in this && !f.call(t, this[i], i, this)) return false; return true; },
      some(f, t) { for (let i = 0; i < this.length; i++) if (i in this && f.call(t, this[i], i, this)) return true; return false; },
      reduce(f, ini) { let i = 0, acc; if (arguments.length > 1) acc = ini; else { while (i < this.length && !(i in this)) i++; if (i >= this.length) throw new TypeError('Reduce of empty array with no initial value'); acc = this[i++]; } for (; i < this.length; i++) if (i in this) acc = f(acc, this[i], i, this); return acc; }
    }],
    [String.prototype, {
      startsWith(x, pos) { const t = String(this), q = String(x), p = Math.max(0, pos | 0); return t.substr(p, q.length) === q; },
      endsWith(x, fin) { const t = String(this), q = String(x), e = fin === undefined ? t.length : Math.min(Math.max(0, fin | 0), t.length); return e - q.length >= 0 && t.substr(e - q.length, q.length) === q; }
    }],
    [Function.prototype, {
      bind(ctx, ...pre) { const f = this; if (typeof f !== 'function') throw new TypeError('Bind must be called on a function'); const B = function (...post) { return new.target ? Reflect.construct(f, pre.concat(post), new.target === B ? f : new.target) : f.apply(ctx, pre.concat(post)); }; return B; }
    }],
    [Object, {
      keys(o) { const x = Object(o), r = []; for (const k of Reflect.ownKeys(x)) if (typeof k === 'string' && Object.prototype.propertyIsEnumerable.call(x, k)) r.push(k); return r; },
      values(o) { const x = Object(o), r = []; for (const k of Reflect.ownKeys(x)) if (typeof k === 'string' && Object.prototype.propertyIsEnumerable.call(x, k)) r.push(x[k]); return r; }
    }],
    [Array, {
      from(it, f, t) { if (it == null) throw new TypeError('Array.from requiere un objeto iterable'); const r = []; if (typeof it[Symbol.iterator] === 'function') { let i = 0; for (const x of it) r.push(f ? f.call(t, x, i++) : x); } else { const n = Math.max(0, Math.min(Number(it.length) || 0, 4294967295)); for (let i = 0; i < n; i++) r.push(f ? f.call(t, it[i], i) : it[i]); } return r; }
    }]
  ];
  let limpioN = 0, limpioHechos = [];
  function limpioEntrar() {
    if (limpioN++ > 0) return;                         // ya hay otra tarea en curso: se comparte el mismo estado
    limpioHechos = [];
    [Array, Object, String].forEach(C => { for (const k in C.prototype) { const d = Object.getOwnPropertyDescriptor(C.prototype, k); if (d && d.enumerable && d.configurable) { Object.defineProperty(C.prototype, k, Object.assign({}, d, { enumerable: false })); limpioHechos.push(['enum', C.prototype, k]); } } });
    estandar().forEach(([obj, mets]) => { Object.getOwnPropertyNames(mets).forEach(k => { const d = Object.getOwnPropertyDescriptor(obj, k); if (d && d.configurable && !esNativa(d.value)) { limpioHechos.push(['est', obj, k, d]); Object.defineProperty(obj, k, { value: mets[k], writable: true, configurable: true, enumerable: false }); } }); });
  }
  function limpioSalir() {
    if (--limpioN > 0) return;
    limpioHechos.forEach(h => { try { if (h[0] === 'est') Object.defineProperty(h[1], h[2], h[3]); else { const d = Object.getOwnPropertyDescriptor(h[1], h[2]); if (d) Object.defineProperty(h[1], h[2], Object.assign({}, d, { enumerable: true })); } } catch (e) { /* se deja como esté */ } });
    limpioHechos = [];
  }
  async function limpio(fn) { limpioEntrar(); try { return await fn(); } finally { limpioSalir(); } }
  const hasGM = typeof GM_getValue === 'function';
  const GMX = {
    get(k, d) { try { const v = hasGM ? GM_getValue(k, null) : localStorage.getItem('pnf.gm.' + k); return v == null ? d : (typeof v === 'string' ? JSON.parse(v) : v); } catch (e) { return d; } },
    set(k, v) { const s = jsonStr(v); try { hasGM ? GM_setValue(k, s) : localStorage.setItem('pnf.gm.' + k, s); } catch (e) { console.warn('PNF: no se pudo guardar', k, e); } }
  };
  function http(o) {
    return new Promise((res, rej) => {
      if (typeof GM_xmlhttpRequest === 'function') GM_xmlhttpRequest({ method: o.method || 'GET', url: o.url, headers: o.headers, data: o.data, timeout: 20000, onload: r => res({ status: r.status, text: r.responseText }), onerror: () => rej(new Error('sin conexión')), ontimeout: () => rej(new Error('tiempo agotado')) });
      else if (W.__PNF_HTTP) W.__PNF_HTTP(o).then(res, rej);
      else fetch(o.url, { method: o.method || 'GET', headers: o.headers, body: o.data }).then(async r => res({ status: r.status, text: await r.text() })).catch(rej);
    });
  }
  const clone = o => JSON.parse(jsonStr(o));
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const pad = n => String(n).padStart(2, '0');
  const hh = (d = new Date()) => pad(d.getHours()) + ':' + pad(d.getMinutes());
  const hoyTxt = () => { const d = new Date(); return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear(); };
  const cuando = iso => { if (!iso) return '—'; const d = new Date(iso); if (isNaN(d)) return '—'; return (d.toDateString() === new Date().toDateString() ? 'hoy ' : pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + ' ') + hh(d); };
  const horasDesde = iso => iso ? (Date.now() - new Date(iso).getTime()) / 36e5 : Infinity;
  const stamp = () => { const d = new Date(); return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()); };
  const fnv = s => { let h = 0x811c9dc5; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; } return h >>> 0; };
  const SESION = Math.random().toString(36).slice(2, 10);
  const setPath = (o, p, v) => { const k = p.split('.'); let x = o; while (k.length > 1) x = x[k.shift()]; x[k[0]] = v; };
  const cmpVer = (a, b) => { const x = a.split('.').map(Number), y = b.split('.').map(Number); for (let i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0) ? 1 : -1; } return 0; };
  // PIN de las secciones protegidas: solo guarda una huella (protección cosmética, no seguridad real)
  const PIN_H = 3567839817;
  const pinOk = v => fnv('pnf|' + String(v).trim()) === PIN_H;

  /* ═══════════════ 2 · Valores iniciales (todo se puede cambiar en Configuración) ═══════════════ */
  const LOGO = 'data:image/png;base64,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Xv3AsflbFrblCzrjNLPUfHFg/ppckFARNQVadigpEqyINcQIMYEv+aWh7DXF0kIwkn3NrAyCkEx2LNA19YgR0AGySYdeNP37N5+xeWH2sC/7rr4+Q0gENjAig4kRQtS2qQS3XLLx/VdvenG49t2HDp23Y+u3739h98FDmChwhBzYID0QAYsqiCaYVgCaJcWMAEbg7BzErzR1owGyjQqBG1bcYVu17z5UX75g0w1X8ZvOKd75/OBnP7+yYYORK/7H+417ffiD/TPv+wMjSAK473/7fDaZpfPHvnCrG0dUiuZdfHbnFaujqux/+7u90HKlHT/fH2zZmuxeEOV1vTLVHBqBzadZAqSJAAQaSxqJCjEGEBjYpmkb4tfgGAsSICCIwIASLhgLgBhNVSXQdVHXbBNhbfjac5a+45KNg/0ng71H91enL77lepiN3XEMEDOyMdjXnjUIaNR5vYnE1raRuHHhytTu3ZbjoqmPtebairUgDA14AthGNMwINBuzodllCL0iSARkKPhkGTQuGx1L4UCp2P+nf9v7bw8k33rF0uuuOXj8yOSBwZQtxu942jm+O3Xcty2pmG2A3mRPsG35vqcPZP76HtErORnW9xwY3nNfMir0TI3LrrxJ+pnju53pS2pLV4leYe8ajAemAOwWTjsBAXBsYWRJC9FjlhoJjCYmja81lggsAMEIxSCFIYsoiKpph5ct7dq177CbLTBro9S6Fd05NKNh7eJty44PpGdqTZn0bCQkDowky0MQJ05O5FNOW87msL65L7HBLliqYUEnTIyu64w+8tvn7jpR/Pb9Lw6PadQpXzRIWMogI7GZZcavtghGADIQCbK0zoYUcjS4KFOz3ey9j6q90+aC86xV88Kjx43RE3c8ynffrubP79EqFwEaU/zo3yR/69qFp58jwZkyfjlv8PnD9sf+zXYXelZ9Ym27sE1zerR733F99hbKoP25uyBhgREBpYQwxoSBpaqOdiVZIVqMlgbJ5lSo8rVll+ZUZNKAIHA4jrqsYPPylo+9dcuX7+O/vfMQdixli58fiCXSovbWbz1//LxztguXar6KKjP51nzDso5P11ukEJlEVXGbSCCR68gY5K69/fHkzOY1rZ/9je0LF7irlizYtHHB+z/0jVqxSakmCyG9TMwSkACQmRjMK4JkgIoLiOBqzEU0HNbbr792w2UXFr/2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';
  const CAT0 = [
    { id: 'paradero', t: 'Procedería infracción, por ser una zona con prohibición expresa de tomar o dejar pasajeros', kw: ['paradero no autorizado'] },
    { id: 'trazado', t: 'Procedería infracción, por cuanto no cumple trazado registrado', kw: ['no dando cumplimiento a su recorrido'] },
    { id: 'revisar', t: '', kw: ['cancelado temporalmente'], dejar: true },
    { id: 'novehic', t: 'No se detectaron vehículos', kw: ['sin ppu', 'no se observan', 'no fue posible constatar', 'no logrando observar'] },
    { id: 'noinf', t: 'No se indica alguna infracción', kw: ['totalidad de la documentación', 'cumpliendo con', 'se realiza recorrido', 'efectúa recorrido por'] },
    { id: 'insumo', t: 'No contamos con el insumo para definir si corresponde o no', kw: [], porDefecto: 'INF' }
  ];
  const CFG0 = {
    de: { nombre: 'MARCELA CANALES INFANTE', cargo: 'ENCARGADA ÁREA DE OPERACIONES - SEREMITT.' },
    para: { nombre: 'SERGIO STEPHAN ORELLANA', cargo: 'JEFE ÁREA LEGAL - SEREMITT.' },
    asunto: 'Remite Análisis Técnico {{TIPO}} N° {{NUMERO}}',
    intro: 'Mediante el presente, remito a usted, análisis técnico referente a {{DOC}}, de la División de Fiscalización, según el siguiente detalle:',
    cierre: 'Saludos cordiales.',
    personas: [
      { usuario: 'rgarcia', nombre: 'RODRIGO ANDRÉS GARCÍA', cargo: 'ENCARGADO/A UNIDAD DE ANÁLISIS', ok: true },
      { usuario: 'mcanales', nombre: 'MARCELA BELÉN CANALES INFANTE', cargo: 'ENCARGADO(A) DE OPERACIONES', ok: true }
    ],
    visador: 'rgarcia', firmante: 'mcanales',
    form: { tipoResp: 'Documento para firma (sólo PDF)', tipoDoc: 'Memorándum', tipoMateria: 'Sin Materia', ciudad: 'Santiago', actos: 'No' },
    listaDist: 'PNF fiscalizacion', pausaGuardar: false
  };
  const datosNuevos = () => ({ v: 1, rev: 0, guardado: '', titular: '', ses: '', oficios: {}, catalogo: clone(CAT0), cfg: clone(CFG0) });
  const rolesNuevos = () => ({ v: 1, rev: 1, titular: 'aestaym', base: 'aestaym', desde: hoyTxt(), por: 'ggodoyb', hasta: '', admins: ['ggodoyb'],
    roles: { aestaym: 'Analista', ggodoyb: 'Administrador', mcanales: 'Supervisor', rgarcia: 'Supervisor' },
    nombres: { aestaym: 'ALEJANDRA ESTAY', ggodoyb: 'GONZALO ANDRÉS GODOY', mcanales: 'MARCELA BELÉN CANALES INFANTE', rgarcia: 'RODRIGO ANDRÉS GARCÍA NEIRA' }, dias: 7, horas: 24, seg: 60,
    hist: [{ f: hoyTxt() + ' ' + hh(), a: 'aestaym', por: 'ggodoyb', motivo: 'Asignación inicial', hasta: 'manual' }] });
  const ROLES_TXT = ['Analista', 'Supervisor', 'Administrador', 'Espectador'];

  /* ═══════════════ 3 · Estado ═══════════════ */
  const L = Object.assign({ usuario: '', nombreSGD: '', bienvenida: false, clave: '', motor: 'gemini', modelo: 'gemini-3.5-flash-lite', fab: null, min: true, visto: 0, ultResp: '', ultCopia: 0 }, GMX.get('pnf.local', {}));
  let R = normR(GMX.get('pnf.roles', null) || rolesNuevos()); R.nombres = R.nombres || {};
  let D = normD(Object.assign(datosNuevos(), GMX.get('pnf.datos', {})));      // datos propios (los que escribe el titular)
  let V = normD(Object.assign(datosNuevos(), GMX.get('pnf.vista', {})));      // última copia leída de la carpeta (para quien no es titular)
  const S = { vista: 'oficios', cfgTab: 'cat', pinOk: false, pinTs: 0, dirty: false, msg: 'Listo', carpeta: 'sin', carpetaNombre: '', ver: null, conflicto: false, errSync: '', sync: false, ultSync: 0, cand: [], uFil: '', ofrecer: null, avisoCerrado: {} };
  const saveL = () => GMX.set('pnf.local', L);
  const datos = () => esTit() ? D : V;
  const rolDe = u => (R.roles && R.roles[u]) || 'Espectador';
  const esTit = () => !!L.usuario && L.usuario === R.titular;
  const esAdm = () => !!L.usuario && (R.admins || []).includes(L.usuario);
  const etiquetaRol = () => !L.usuario ? 'sin identificar' : esTit() ? 'Titular' + (esAdm() ? ' · Administrador' : '') : (esAdm() ? 'Administrador' : rolDe(L.usuario));
  const guardaPIN = () => { S.pinTs = Date.now(); };
  const pinVigente = () => S.pinOk && (Date.now() - S.pinTs < 10 * 60 * 1000);

  /* ═══════════════ 4 · Carpeta de respaldo (File System Access) ═══════════════ */
  const FS = {
    _db: null, h: null,
    abrir() { return new Promise((res, rej) => { if (this._db) return res(this._db); const q = indexedDB.open('pnf-carpeta', 1); q.onupgradeneeded = () => q.result.createObjectStore('k'); q.onsuccess = () => { this._db = q.result; res(q.result); }; q.onerror = () => rej(q.error); }); },
    async get(k) { const db = await this.abrir(); return new Promise((res, rej) => { const q = db.transaction('k').objectStore('k').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
    async set(k, v) { const db = await this.abrir(); return new Promise((res, rej) => { const q = db.transaction('k', 'readwrite').objectStore('k').put(v, k); q.onsuccess = () => res(); q.onerror = () => rej(q.error); }); },
    async cargarGuardada() { try { this.h = await this.get('carpeta') || null; } catch (e) { this.h = null; } return this.h; },
    async elegir() { const f = W.showDirectoryPicker || window.showDirectoryPicker; if (!f) throw new Error('Este navegador no permite elegir carpetas'); this.h = await f.call(W, { mode: 'readwrite', id: 'pnf-respaldo' }); await this.set('carpeta', this.h); return this.h; },
    async permiso(pedir) { if (!this.h) return false; try { let p = await this.h.queryPermission({ mode: 'readwrite' }); if (p !== 'granted' && pedir) p = await this.h.requestPermission({ mode: 'readwrite' }); return p === 'granted'; } catch (e) { return false; } },
    async leer(nombre, sub) { let d = this.h; try { if (sub) d = await d.getDirectoryHandle(sub); const f = await (await d.getFileHandle(nombre)).getFile(); return await f.text(); } catch (e) { if (e && (e.name === 'NotFoundError' || e.name === 'TypeMismatchError')) return null; throw e; } },
    async escribir(nombre, texto, sub) { let d = this.h; if (sub) d = await d.getDirectoryHandle(sub, { create: true }); const w = await (await d.getFileHandle(nombre, { create: true })).createWritable(); await w.write(texto); await w.close(); },
    async listar(sub) { const out = []; try { const d = await this.h.getDirectoryHandle(sub); for await (const [n, x] of d.entries()) if (x.kind === 'file') out.push(n); } catch (e) { /* no existe */ } return out; },
    async borrar(nombre, sub) { const d = await this.h.getDirectoryHandle(sub); await d.removeEntry(nombre); }
  };
  async function leerJSON(nombre, sub) { const t = await FS.leer(nombre, sub); return t == null ? null : JSON.parse(t); }
  async function conectar(pedir) {
    if (!FS.h) await FS.cargarGuardada();
    if (!FS.h) { S.carpeta = 'sin'; S.carpetaNombre = ''; return false; }
    S.carpetaNombre = FS.h.name || '';
    const ok = await FS.permiso(!!pedir);
    S.carpeta = ok ? 'ok' : 'permiso';
    return ok;
  }
  async function elegirCarpeta() {
    try { await FS.elegir(); } catch (e) { if (e && e.name === 'AbortError') return; toast('No se pudo elegir la carpeta: ' + e.message); return; }
    S.carpeta = 'ok'; S.carpetaNombre = FS.h.name || '';
    toast('Carpeta conectada: ' + S.carpetaNombre);
    await sincronizar('carpeta'); render();
  }
  const versionArchivo = () => ({ rev: D.rev, ts: D.guardado, titular: R.titular, ses: SESION, v: VERSION });
  let copiaN = 0;
  async function copiaSeguridad() {
    if (S.carpeta !== 'ok') return;
    const texto = jsonStr(D);
    await FS.escribir('PNF-datos-' + stamp() + '-' + String(++copiaN % 1000).padStart(3, '0') + '.json', texto, 'copias');
    const lst = (await FS.listar('copias')).filter(n => /^PNF-datos-.*\.json$/.test(n)).sort();
    while (lst.length > 10) await FS.borrar(lst.shift(), 'copias');
    L.ultCopia = Date.now(); saveL();
  }
  async function escribirDatos(forzarCopia) {
    if (S.carpeta === 'ok') {
      await FS.escribir('PNF-datos.json', jsonStr(D));
      await FS.escribir('PNF-version.json', jsonStr(versionArchivo()));
      if (forzarCopia || Date.now() - L.ultCopia > 10 * 60 * 1000) await copiaSeguridad();
      L.ultResp = new Date().toISOString(); saveL();
      S.ver = versionArchivo();
      return true;
    }
    // sin carpeta: copia en Descargas\PNF-respaldos (si Tampermonkey lo permite)
    try {
      if (typeof GM_download === 'function') { const u = URL.createObjectURL(new Blob([jsonStr(D)], { type: 'application/json' })); GM_download({ url: u, name: 'PNF-respaldos/PNF-datos.json', saveAs: false, conflictAction: 'overwrite' }); L.ultResp = new Date().toISOString(); saveL(); }
    } catch (e) { /* sin descarga */ }
    return false;
  }
  async function escribirRoles() { await FS.escribir('PNF-roles.json', jsonStr(R)); }

  /* ═══════════════ 5 · Guardado, sincronización y alertas ═══════════════ */
  async function guardar(motivo, forzarCopia) {
    if (!esTit()) return false;
    try { await chequearVersion(); } catch (e) { /* sin carpeta: se guarda local */ }
    if (S.conflicto) { toast('⚠ No se guardó: otra sesión guardó antes como titular. Pulsa «↻ Actualizar».'); render(); return false; }
    D.rev++; D.guardado = new Date().toISOString(); D.ses = SESION; D.titular = L.usuario;
    GMX.set('pnf.datos', D); S.dirty = false;
    let enCarpeta = false;
    try { enCarpeta = await escribirDatos(forzarCopia); } catch (e) { S.errSync = 'No se pudo escribir en la carpeta: ' + e.message; }
    toast('Guardado ✓ ' + hh() + (motivo ? ' · ' + motivo : '') + (enCarpeta ? '' : ' (solo en este equipo)'));
    render(); return true;
  }
  const marcar = () => { if (esTit() && !S.dirty) { S.dirty = true; pintarEstado(); } };
  async function chequearVersion() {
    if (S.carpeta !== 'ok') return;
    const v = await leerJSON('PNF-version.json'); S.ver = v;
    S.conflicto = !!(esTit() && v && v.rev > D.rev);
  }
  async function guardarRoles(motivo) {
    R.rev = (R.rev || 0) + 1; GMX.set('pnf.roles', R);
    try { if (S.carpeta === 'ok') await escribirRoles(); toast('Roles guardados ✓ ' + hh() + (motivo ? ' · ' + motivo : '')); } catch (e) { toast('No se pudo escribir PNF-roles.json: ' + e.message); }
    render();
  }
  async function cargarDeCarpeta(motivo) {
    if (S.carpeta !== 'ok') return false;
    const d = await leerJSON('PNF-datos.json'); if (!d) return false;
    try { await copiaSeguridad(); } catch (e) { /* ignorar */ }
    GMX.set('pnf.datos.anterior', D);
    D = normD(Object.assign(datosNuevos(), d)); GMX.set('pnf.datos', D); S.conflicto = false; S.dirty = false; S.ofrecer = null;
    toast('Datos cargados desde la carpeta (revisión ' + D.rev + ')' + (motivo ? ' · ' + motivo : '')); return true;
  }
  async function sincronizar() {
    if (S.sync) return; S.sync = true; S.errSync = '';
    try {
      if (!(await conectar(false))) return;
      const r = await leerJSON('PNF-roles.json');
      if (r === null) { if (esAdm()) await escribirRoles(); }
      else if (r.rev > (R.rev || 0)) { R = normR(r); R.nombres = R.nombres || {}; GMX.set('pnf.roles', R); toast('Se actualizaron los usuarios y roles'); }
      const v = await leerJSON('PNF-version.json'); S.ver = v;
      if (esTit()) {
        S.conflicto = !!(v && v.rev > D.rev);
        if (!v && D.rev > 0) await escribirDatos(true);
        else if (v && v.rev > 0 && D.rev === 0) S.ofrecer = v;
      } else if (v && v.rev > (V.rev || 0)) {
        const d = await leerJSON('PNF-datos.json'); if (d) { V = normD(Object.assign(datosNuevos(), d)); GMX.set('pnf.vista', V); }
      }
      await leerAportes(); if (esTit()) await procesarBuzon();
      S.ultSync = Date.now();
    } catch (e) { S.errSync = e.message || String(e); }
    finally { S.sync = false; }
  }
  function alertas() {
    const A = [], t = esTit(), bt = (x, a, c) => `<button class="vb peq ${c || ''}" ${a}>${x}</button>`;
    if (!L.usuario) A.push(['mal', '⚠ Aún no se sabe quién eres en PNF' + (S.sgd ? ' (entraste al SGD como ' + esc(S.sgd.nombre) + ')' : '') + '. ' + bt('Identificarme', 'data-a="ident"', 'def')]);
    if (S.carpeta === 'sin') A.push(['warn', '📁 La carpeta de respaldo aún no está conectada. Sin ella no hay respaldo automático ni se pueden compartir los datos. ' + bt('Elegir carpeta…', 'data-a="carpeta"', 'def')]);
    if (S.carpeta === 'permiso') A.push(['warn', '📁 Chrome pide confirmar de nuevo el permiso de la carpeta «' + esc(S.carpetaNombre) + '». Mientras tanto no hay respaldo automático. ' + bt('Permitir', 'data-a="permiso"', 'def')]);
    if (S.errSync) A.push(['mal', '⚠ ' + esc(S.errSync)]);
    if (t && S.conflicto) A.push(['mal', '⚠ <b>Otra sesión guardó datos como titular</b> (revisión ' + (S.ver ? S.ver.rev : '?') + ') y tu copia es la ' + D.rev + '. Para no pisar nada, PNF no guardará hasta que actualices. ' + bt('Actualizar y revisar', 'data-a="adoptar"', 'def')]);
    if (t && S.ofrecer) A.push(['inf', 'ℹ En la carpeta hay datos guardados (revisión ' + S.ofrecer.rev + ', ' + cuando(S.ofrecer.ts) + ') y aquí no hay nada. ' + bt('Restaurarlos', 'data-a="adoptar"', 'def')]);
    if (!t && L.usuario) {
      if (S.ver && horasDesde(S.ver.ts) > R.horas) A.push(['mal', '⏱ <b>Datos antiguos:</b> el último guardado de ' + esc(R.titular) + ' fue ' + cuando(S.ver.ts) + ' (límite ' + R.horas + ' h). Puede que no haya trabajado o que Drive no haya traído la última versión. ' + bt('Actualizar ahora', 'data-a="act"', 'def')]);
      if (S.ver && S.ver.rev > (L.visto || 0) && (L.visto || 0) > 0) A.push(['inf', 'ℹ <b>' + esc(R.titular) + '</b> guardó cambios desde tu última vista (revisión ' + L.visto + ' → ' + S.ver.rev + '). ' + bt('Entendido', 'data-a="visto"')]);
    }
    const dt = datos(), venc = Object.keys(dt.oficios || {}).filter(e => dt.oficios[e].sgd !== 'Respondido' && diasDe(dt.oficios[e]) >= R.dias);
    if (venc.length) A.push(['warn', '📅 <b>' + venc.length + ' expediente(s)</b> llevan ' + R.dias + ' días o más sin responder.']);
    Object.keys(dt.oficios || {}).filter(e => dt.oficios[e].sgd === 'Por despachar').forEach(e => A.push(['warn', `📤 <b>${esc(e)}</b> está preparado en el SGD${dt.oficios[e].preparado ? ' (' + esc(cuando(dt.oficios[e].preparado)) + ')' : ''} y <b>falta pulsar «Despachar»</b>. ` + (t ? bt('Ir a despachar', 'data-a="ir-despachar" data-id="' + esc(e) + '"', 'def') + ' ' + bt('Crear el memo de nuevo', 'data-a="crear-nuevo" data-id="' + esc(e) + '"') : '')]));
    alertasAportes(A);
    devueltos().forEach(e => A.push(['mal', `⚠ <b>${esc(e)}</b> volvió a la bandeja: ya lo habías respondido el ${esc(datos().oficios[e].respondido || '—')}, por eso figura como <b>DEVUELTO</b>. ` + (t ? bt('Marcar atendido', 'data-a="atend" data-id="' + esc(e) + '"', 'peq') : '')]));
    if (t) S.aportes.filter(a => a.tipo === 'oficio' && estAp(a.id) === 'incorporado' && !(L.vistos || {})[a.id]).forEach(a => A.push(['inf', `➕ <b>${esc(a.de)}</b> agregó el oficio <b>${esc(a.exp)}</b> (${esc(cuando(a.hora))}). Ya está en tu tabla, marcado con ＋ «Agregado por ${esc(a.de)}». ` + bt('Entendido', 'data-a="vis" data-id="' + esc(a.id) + '"', 'peq')]));
    return A;
  }
  function diasDesde(f) { const p = String(f || '').split('/').map(Number); if (p.length < 3) return 0; return Math.max(0, Math.round((Date.now() - new Date(p[2], p[1] - 1, p[0]).getTime()) / 864e5)); }

  /* ═══════════════ 6 · Quién es la persona (usuario del SGD) ═══════════════ */
  const normNom = x => String(x || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
  // El SGD muestra arriba a la derecha: «Usuario conectado: NOMBRE APELLIDO | Cargo: CARGO | Desconectar» (no muestra el usuario).
  function leerSesionSGD() {
    const t = (document.body ? document.body.innerText : '') || '';
    const m = t.match(/Usuario conectado:\s*([^|\n]+?)\s*\|\s*Cargo:\s*([^|\n]+?)\s*(?:\||\n|$)/i);
    return m ? { nombre: m[1].trim(), cargo: m[2].trim() } : null;
  }
  // El SGD omite el apellido materno: coincide si todas las palabras del nombre más corto están en el más largo (mínimo 2 palabras).
  const mismoNombre = (a, b) => { const x = normNom(a).split(' ').filter(Boolean), y = normNom(b).split(' ').filter(Boolean); if (x.length < 2 || y.length < 2) return false; const c = x.length <= y.length ? x : y, g = x.length <= y.length ? y : x; return c.every(w => g.includes(w)); };
  const candidatosPorNombre = nom => Object.keys(R.roles).filter(u => R.nombres && R.nombres[u] && mismoNombre(R.nombres[u], nom));
  function verificarSesion() {
    const s2 = leerSesionSGD(); if (!s2) return '';
    S.sgd = s2;
    if (L.usuario && L.nombreSGD && !mismoNombre(L.nombreSGD, s2.nombre)) { L.usuario = ''; L.nombreSGD = ''; saveL(); return 'cambio'; }  // entró otra persona al SGD en este navegador
    if (!L.usuario) { const c = candidatosPorNombre(s2.nombre); if (c.length === 1) { L.usuario = c[0]; L.nombreSGD = s2.nombre; saveL(); return 'auto'; } }
    if (L.usuario && !L.nombreSGD) { L.nombreSGD = s2.nombre; saveL(); }
    return '';
  }

  /* ═══════════════ 7 · Interfaz (Shadow DOM sobre el SGD) ═══════════════ */
  const CSS = `
:host { all: initial; }
*{box-sizing:border-box}
.pnf { --face:#e3e1dc; --page:#ebe9e4; --hi:#fff; --sh:#a8a8a8; --title:linear-gradient(90deg,#17387a,#6f9ad0); --font:"Segoe UI",Tahoma,Arial,sans-serif; --r:3px; --rf:5px;
  --btnbg:linear-gradient(#fefefe,#dadada); --btnhov:linear-gradient(#f6fbff,#cfe3f5); --bhi:#8e8e8e; --bdk:#6e6e6e; --hovb:#2c628b; --defb:#2c628b; --ib:#8c8c8c; --ish:inset 0 1px 2px rgba(0,0,0,.2);
  --sel:#2f64b5; --band:#f2f5f9; --grid:#c2c6cc; --thbg:linear-gradient(#fff,#d9d9d9); --fb:#9a9a9a; --fbg:#f7f6f3; --legend:#1d3d6e; --tabbg:linear-gradient(#f4f3ef,#d9d7d1); --acc:#2f64b5;
  --ok:#d6f0cf; --warn:#fff5c2; --bad:#f8d2d2; --note:#fff8d0; font:12px var(--font); color:#000; }
.win { position:fixed; top:14px; left:50%; transform:translateX(-50%); width:min(1280px,calc(100vw - 28px)); height:calc(100vh - 28px); display:none; flex-direction:column; z-index:2147483000; background:var(--face); border:1px solid #6f8fb5; border-radius:5px; box-shadow:0 8px 26px rgba(0,0,0,.45); padding:3px; }
.win.on { display:flex; }
.titlebar { background:var(--title); color:#fff; font-weight:bold; font-size:13px; display:flex; align-items:center; gap:6px; border-radius:3px 3px 0 0; margin:-3px -3px 3px; padding:4px 6px; flex:none; }
.titlebar .ico { width:16px; height:16px; background:#fff; color:#17387a; display:inline-flex; align-items:center; justify-content:center; font-size:11px; border-radius:3px; }
.sp { flex:1; }
.cap { width:19px; height:17px; background:linear-gradient(#fff,#d4d4d4); color:#000; font-size:11px; line-height:13px; text-align:center; border:1px solid #8a8a8a; border-radius:3px; font-weight:bold; cursor:pointer; }
.cap:hover { background:linear-gradient(#f6fbff,#cfe3f5); } .cap.x:hover { background:linear-gradient(#f59a8a,#d6452b); color:#fff; }
.toolbar { display:flex; align-items:center; gap:3px; padding:3px 4px; background:linear-gradient(#fbfbfb,#dedcd6); border-bottom:1px solid var(--sh); flex:none; }
.toolbar .tb { font:inherit; padding:2px 9px; background:transparent; border:1px solid transparent; border-radius:3px; cursor:pointer; }
.toolbar .tb:hover { border-color:var(--hovb); background:var(--btnhov); } .toolbar .tb:disabled { opacity:.5; cursor:not-allowed; }
.tsep { width:1px; height:20px; background:var(--sh); margin:0 4px; }
.form { flex:1; overflow:auto; padding:8px 6px 6px; min-height:0; }
.status { display:flex; gap:3px; padding:3px 3px 0; flex:none; }
.status span { border:1px solid; border-color:var(--sh) #fff #fff var(--sh); padding:1px 8px; white-space:nowrap; border-radius:3px; }
.status span:first-child { flex:1; overflow:hidden; text-overflow:ellipsis; } .status span.dirty { background:#ffe08a; font-weight:bold; }
.stbtn { font:inherit; font-size:11px; padding:0 6px; border:1px solid var(--sh); background:linear-gradient(#fff,#e1e1e1); border-radius:3px; cursor:pointer; margin-left:6px; }
.mtabs { display:flex; padding-left:4px; position:relative; z-index:1; flex-wrap:wrap; }
.mtab { font:inherit; padding:3px 13px; background:var(--tabbg); border:1px solid; border-color:#8a8a8a #8a8a8a transparent #8a8a8a; border-radius:5px 5px 0 0; margin-top:2px; position:relative; top:1px; color:#000; cursor:pointer; }
.mtab:hover { background:linear-gradient(#f6fbff,#d3e5f5); }
.mtab.on { margin-top:0; padding-top:5px; padding-bottom:4px; font-weight:bold; z-index:2; background:var(--page); box-shadow:inset 0 3px 0 var(--acc); }
.mtab.fut { color:#777; }
.mtabs .sepx { flex:1; }
.mpage { border:1px solid #8a8a8a; box-shadow:inset 0 1px 0 #fff; padding:10px; background:var(--page); border-radius:0 0 5px 5px; min-height:420px; }
.vb { font:inherit; color:#000; padding:3px 13px; min-width:76px; cursor:pointer; background:var(--btnbg); border:1px solid; border-color:var(--bhi) var(--bdk) var(--bdk) var(--bhi); border-radius:3px; box-shadow:inset 0 1px 0 #fff; outline:0; }
.vb:hover { background:var(--btnhov); border-color:var(--hovb); } .vb:active { border-color:var(--bdk) var(--bhi) var(--bhi) var(--bdk); }
.vb.def { border:2px solid var(--defb); padding:2px 12px; font-weight:bold; } .vb.peq { min-width:0; padding:1px 9px; } .vb.peq.def { padding:0 8px; }
.vb[disabled] { color:#8a8a8a; cursor:not-allowed; opacity:.7; }
fieldset.fr { border:1px solid var(--fb); border-radius:5px; box-shadow:0 1px 0 #fff, inset 0 0 0 1px #fff; background:var(--fbg); margin:0 0 10px; padding:6px 10px 9px; min-width:0; }
fieldset.fr legend { padding:0 5px; color:var(--legend); font-weight:bold; }
input[type=text], input[type=password], input[type=number], select, textarea { font:inherit; background:#fff; color:#000; border:1px solid var(--ib); box-shadow:var(--ish); padding:2px 4px; border-radius:3px; max-width:100%; }
input:focus, select:focus, textarea:focus { outline:0; border-color:var(--acc); box-shadow:0 0 0 1px var(--acc); }
textarea { width:100%; min-height:54px; resize:vertical; } input[type=checkbox], input[type=radio] { accent-color:var(--acc); }
label.lb { display:block; margin:6px 0 2px; }
.row { display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
.g2 { display:grid; grid-template-columns:1fr 1fr; gap:10px; } @media (max-width:1000px){ .g2{grid-template-columns:1fr;} }
.estado { display:inline-block; padding:0 7px; border:1px solid var(--sh); background:#fff; border-radius:3px; }
.e-z { background:#d8e8ff; } .e-p { background:#e6dcff; } .e-v { background:var(--ok); } .e-a { background:var(--warn); } .e-r { background:var(--bad); }
.kv { display:grid; grid-template-columns:130px minmax(0,1fr); gap:4px 10px; margin:0; } .kv dt { color:#444; } .kv dd { margin:0; min-width:0; word-break:break-word; }
.note { background:var(--note); border:1px solid var(--fb); border-radius:3px; padding:4px 9px; margin-bottom:8px; }
.note.ok { background:var(--ok); border-color:#7fb77a; } .note.mal { background:var(--bad); border-color:#c98585; } .note.inf { background:#e3effa; border-color:#8fb3d6; } .note.warn { background:var(--warn); border-color:#d6c46a; }
.ro { background:#eef3f8; border:1px solid #b5c6da; border-radius:3px; padding:3px 9px; margin-bottom:8px; font-size:11.5px; }
.rolchip { display:inline-block; padding:0 7px; border:1px solid #6f8fb8; background:#e3effa; border-radius:3px; }
.mono { font-family:"Courier New",monospace; font-size:12px; }
.lv { background:#fff; border:1px solid var(--ib); box-shadow:var(--ish); overflow:auto; border-radius:3px; }
.lv table { border-collapse:collapse; width:100%; }
.lv th { background:var(--thbg); text-align:left; font-weight:normal; padding:3px 7px; border:1px solid; border-color:#fff #8a8a8a #8a8a8a #fff; position:sticky; top:0; white-space:nowrap; z-index:1; }
.lv td { padding:3px 7px; border:1px solid var(--grid); vertical-align:top; } .lv tr:nth-child(even) td { background:var(--band); }
.pinbox { max-width:440px; margin:26px auto; text-align:center; padding:16px; border:1px solid var(--fb); background:var(--note); border-radius:3px; box-shadow:0 1px 3px rgba(0,0,0,.2); }
.vacio { text-align:center; padding:26px 10px; color:#555; }
.pdfi { display:inline-block; position:relative; width:26px; height:30px; background:#d32f2f; border-radius:3px; color:#fff; font:800 9px Arial,sans-serif; text-align:center; line-height:36px; cursor:pointer; vertical-align:middle; margin:0 6px; box-shadow:0 1px 3px rgba(0,0,0,.4); }
.pdfi::before { content:''; position:absolute; top:0; right:0; width:9px; height:9px; background:linear-gradient(225deg,#f7f6f3 50%,#9c1c1c 50%); border-top-right-radius:3px; }
.pdfi:hover { background:#e53935; }
.busy { display:none; position:absolute; left:0; right:0; bottom:0; top:30px; background:rgba(235,233,228,.85); z-index:20; align-items:center; justify-content:center; }
.busy.on { display:flex; }
.busy .bx { background:#fff; border:1px solid #6f8fb5; border-radius:6px; padding:22px 34px; text-align:center; box-shadow:0 6px 20px rgba(0,0,0,.35); min-width:340px; max-width:540px; }
.busy .spin { width:46px; height:46px; margin:0 auto 12px; border:5px solid #cfdcec; border-top-color:#17387a; border-radius:50%; animation:pnfgira .9s linear infinite; }
.busy .t1 { font-size:15px; font-weight:bold; color:#17387a; } .busy .t2 { margin-top:4px; word-break:break-word; } .busy .t3 { margin-top:10px; font-size:11.5px; color:#555; } .busy .t4 { margin-top:6px; font-size:12px; color:#900; font-weight:bold; }
@keyframes pnfgira { to { transform:rotate(360deg); } }
.ov { position:fixed; inset:0; background:rgba(0,0,0,.25); display:none; align-items:center; justify-content:center; z-index:2147483200; } .ov.on { display:flex; }
.mb { background:var(--face); border:1px solid #6f8fb5; border-radius:5px; box-shadow:0 8px 26px rgba(0,0,0,.45); padding:3px; min-width:340px; width:560px; max-width:94vw; }
.mb .tbar { background:var(--title); color:#fff; font-weight:bold; padding:3px 6px; border-radius:3px 3px 0 0; margin:-3px -3px 0; }
.mb .bd { padding:14px 18px; max-height:70vh; overflow:auto; } .mb .bt { text-align:center; padding:0 0 10px; display:flex; gap:8px; justify-content:center; flex-wrap:wrap; }
.mb .err { color:#b00020; min-height:16px; padding:0 18px; }
#tip { position:fixed; z-index:2147483400; background:#ffffe1; border:1px solid #555; box-shadow:2px 2px 6px rgba(0,0,0,.3); padding:5px 8px; max-width:520px; display:none; pointer-events:none; white-space:pre-wrap; font-size:12px; color:#000; }
#fab { position:fixed; right:18px; bottom:18px; width:58px; height:58px; border-radius:50%; z-index:2147483100; display:none; flex-direction:column; align-items:center; justify-content:center; cursor:grab; user-select:none; touch-action:none; border:3px solid #fff; box-shadow:0 4px 12px rgba(0,0,0,.45); background:linear-gradient(180deg,#0063AF 0 50%,#E72E39 50% 100%); }
#fab.on { display:flex; } #fab:active { cursor:grabbing; }
#fab b { font:800 17px Arial,Helvetica,sans-serif; color:#fff; line-height:1; text-shadow:-1px -1px 0 #3a0a10,1px -1px 0 #3a0a10,-1px 1px 0 #3a0a10,1px 1px 0 #3a0a10,0 2px 3px rgba(0,0,0,.4); }
#fab small { font:700 8.5px Arial,sans-serif; color:#fff; line-height:1; margin-top:2px; text-shadow:-.5px -.5px 0 #3a0a10,.5px -.5px 0 #3a0a10,-.5px .5px 0 #3a0a10,.5px .5px 0 #3a0a10; }
#fab .bdg { position:absolute; top:-7px; right:-7px; min-width:21px; height:21px; border-radius:11px; background:#f08a1c; color:#fff; font:700 11px Arial; border:2px solid #fff; display:none; align-items:center; justify-content:center; }
#fab .bdg.on { display:flex; }
#fab .bdg3 { top:-7px; right:auto; left:-7px; background:#e65100; }
#fab #fabN { background:#d32f2f; } #fab .bdg2 { top:auto; bottom:-7px; background:#2e7d32; }
#card { position:fixed; right:16px; bottom:16px; width:400px; z-index:2147483300; background:var(--face); border:1px solid #6f8fb5; border-radius:5px; box-shadow:0 8px 28px rgba(0,0,0,.45); display:none; padding:3px; }
#card.on { display:block; } #card .h { background:var(--title); color:#fff; padding:4px 8px; font-weight:bold; border-radius:3px 3px 0 0; margin:-3px -3px 3px; }
#card .b { padding:6px 8px 8px; max-height:460px; overflow:auto; background:var(--fbg); border:1px solid var(--fb); border-radius:3px; } #card .p { padding:2px 0; } #card .p.ok { color:#1b6e1b; }

.an { position:relative; background:#fff; border:1px solid var(--ib); box-shadow:var(--ish); border-radius:3px; padding:3px 22px 3px 6px; min-height:24px; } .an::after { content:'▾'; position:absolute; right:7px; top:3px; color:#333; }
.an select { position:absolute; inset:0; width:100%; height:100%; opacity:0; cursor:pointer; } .an:hover { border-color:var(--hovb); }
td.sug { background:var(--warn) !important; } td.okc { background:var(--ok) !important; } td.sinc { background:var(--bad) !important; }
.lv tr.foco td { box-shadow:inset 0 2px 0 var(--acc), inset 0 -2px 0 var(--acc); }
.txt { max-width:330px; } .txt.ancho { max-width:560px; min-width:320px; }
.pb { height:16px; background:#fff; border:1px solid var(--ib); box-shadow:var(--ish); padding:1px; min-width:240px; border-radius:3px; } .pb i { display:block; height:100%; background:linear-gradient(#93e093,#2d9a2d); border-radius:2px; }
.visor { background:#808080; padding:14px; overflow:auto; max-height:640px; border:1px solid var(--ib); box-shadow:var(--ish); }
.xl td:first-child, .xl th:first-child { background:#d8e8ff; font-weight:bold; position:sticky; left:0; }
.sunk { background:#fff; border:1px solid var(--ib); box-shadow:var(--ish); border-radius:3px; }

.tw { background:#fff; border:1px solid var(--ib); box-shadow:var(--ish); border-radius:3px; }
table.t { border-collapse:collapse; width:100%; table-layout:fixed; }
table.t th { background:var(--thbg); text-align:left; font-weight:normal; padding:3px 6px; border:1px solid; border-color:#fff #8a8a8a #8a8a8a #fff; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
table.t td { padding:3px 6px; border:1px solid var(--grid); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; height:25px; }
table.t tbody tr:nth-child(even) td { background:var(--band); } table.t tbody tr:hover td { filter:brightness(.97); cursor:default; }
table.t tr.sel td { background:var(--sel) !important; color:#fff; } table.t tr.sel .estado { color:#000; } table.t td.c { text-align:center; } table.t tr.sel .exp { color:#fff; }
table.t tr.dv td:first-child { box-shadow:inset 3px 0 0 #f08a1c; }
table.t tr.fl th { background:#f4f3ef; padding:2px 3px; border-color:#ddd #bbb #bbb #ddd; } table.t tr.fl input, table.t tr.fl select { width:100%; font-size:11px; padding:1px 3px; }
.exp { font-weight:bold; font-family:"Courier New",monospace; color:#17387a; } .lv .exp { white-space:nowrap; }
.tag { font-size:10.5px; color:#2f64b5; font-weight:bold; }
.ord { cursor:pointer; } .ord:hover { text-decoration:underline; } .sortk { color:#2f64b5; font-size:10px; margin-left:3px; }
.pbm { display:inline-block; width:46px; height:10px; background:#fff; border:1px solid var(--ib); border-radius:2px; vertical-align:middle; overflow:hidden; } .pbm i { display:block; height:100%; background:linear-gradient(#93e093,#2d9a2d); }
.pag { display:flex; gap:10px; align-items:center; padding:7px 2px 0; flex-wrap:wrap; }
.pag .n { display:flex; gap:3px; } .pag .n span { min-width:25px; text-align:center; padding:2px 6px; border:1px solid #8e8e8e; border-radius:3px; background:var(--btnbg); cursor:pointer; } .pag .n span:hover { background:var(--btnhov); }
.pag .n span.on { background:var(--sel); color:#fff; border-color:var(--sel); } .pag .n span.dis { color:#999; cursor:default; background:#eee; }
.drop { border:2px dashed #8a8a8a; background:#fff; border-radius:5px; padding:16px; text-align:center; cursor:pointer; } .drop:hover, .drop.sobre { background:#eaf3fc; border-color:var(--acc); }
.cb { position:relative; display:inline-block; }
.cbList { display:none; position:absolute; left:0; top:100%; z-index:60; width:540px; max-width:80vw; max-height:310px; overflow:auto; background:#fff; border:1px solid var(--ib); box-shadow:0 4px 12px rgba(0,0,0,.35); border-radius:3px; } .cbList.on { display:block; }
.cbi { padding:3px 8px; cursor:pointer; border-bottom:1px solid var(--grid); } .cbi.hl { background:var(--sel); color:#fff; } .cbi.hl .estado { color:#000; }
.cbs { font-size:11px; opacity:.85; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; } .cbf { padding:3px 8px; font-size:11px; color:#555; background:#f4f4f4; }
`;

  const host = document.createElement('div'); host.id = 'pnf-host';
  host.style.cssText = 'all:initial;position:fixed;z-index:2147483000;top:0;left:0;';
  document.documentElement.appendChild(host);
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `<style>${CSS}</style><div class="pnf">
  <div class="win" id="win">
    <div class="titlebar"><span class="ico">P</span><span>PNF — Respuesta a oficios de fiscalización</span><span class="sp"></span><span class="cap" id="bMin" title="Minimizar">&#8211;</span><span class="cap x" id="bCerrar" title="Cerrar (queda el ícono)">✕</span></div>
    <div class="toolbar" id="toolbar"></div>
    <div class="form"><div class="mtabs" id="tabs"></div><div class="mpage" id="vista"></div></div>
    <div class="status"><span id="st1">Listo</span><span id="st2"></span><span id="st3"></span><span id="stG"></span><span id="stV"></span></div>
    <div class="busy" id="busy"><div class="bx"><div class="spin"></div><div class="t1" id="busyT"></div><div class="t2" id="busyS"></div><div class="t3" id="busyN"></div><div class="t4">No cambies de pantalla ni cierres PNF hasta que termine.</div></div></div>
  </div>
  <div id="tip"></div>
  <div id="card"><div class="h"><span id="cardT"></span></div><div class="b" id="cardB"></div></div><input type="file" id="fileIn" accept="application/pdf,.pdf" multiple style="display:none"><div id="fab" title="Abrir PNF · arrastra para mover"><b>PNF</b><small>ggb</small><span class="bdg" id="fabN">0</span><span class="bdg bdg2" id="fabP">0</span><span class="bdg bdg3" id="fabD">⚠</span></div>
  <div class="ov" id="pdfOv"><div class="mb" style="width:900px;max-width:96vw"><div class="tbar" style="display:flex;justify-content:space-between;align-items:center"><span id="pdfT">Vista previa</span><span class="cap x" data-a="pdf-cerrar" title="Cerrar (Esc)">✕</span></div><div style="padding:8px 4px 4px"><div class="visor pnfmemo" id="pdfBody" style="max-height:72vh"></div></div><div class="bt" style="padding:6px 0 10px"><button class="vb def" data-a="pdf-cerrar">Cerrar</button></div></div></div>
  <div class="ov" id="ov"><div class="mb" id="mb"><div class="tbar" id="mbT">PNF</div><div class="bd" id="mbB"></div><div class="err" id="mbE"></div><div class="bt" id="mbBt"></div></div></div>
  </div>`;
  const $ = s => root.querySelector(s);
  const $$ = s => [...root.querySelectorAll(s)];

  // vista previa del memo (la misma hoja que se adjunta como PDF)
  function abrirPdf() { const ex = exAct(), o = (datos().oficios || {})[ex]; if (!o) return; $('#pdfT').textContent = 'Vista previa · ' + nombreMemo(o) + '.pdf'; $('#pdfBody').innerHTML = `<style>${MEMO_CSS}</style>` + hojaMemo(ex, false); $('#pdfBody').scrollTop = 0; $('#pdfOv').classList.add('on'); }
  function cerrarPdf() { $('#pdfOv').classList.remove('on'); }
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#pdfOv').classList.contains('on')) { e.stopPropagation(); cerrarPdf(); } }, true);
  // aviso con movimiento mientras PNF trabaja: tapa la ventana (menos la barra de título) para que nadie haga clic en otro lugar
  let busyIv = null;
  function ocupado(titulo, detalle) {
    $('#busyT').textContent = titulo; $('#busyS').textContent = detalle || ''; $('#busy').classList.add('on');
    const t0 = Date.now(), n = $('#busyN'); n.textContent = 'Trabajando… 0 s'; clearInterval(busyIv);
    busyIv = setInterval(() => { n.textContent = 'Trabajando… ' + Math.round((Date.now() - t0) / 1000) + ' s'; }, 1000);
  }
  function libre() { clearInterval(busyIv); busyIv = null; const b = $('#busy'); if (b) b.classList.remove('on'); }
  const toast = m => { S.msg = m; const e = $('#st1'); if (e) e.textContent = m; };
  let dlgCierra = null;
  function dialogo(titulo, html, botones, validar) {
    if (dlgCierra) dlgCierra('cancel');
    libre();                                  // si hay un cuadro de aviso, se quita el «trabajando» (el cuadro lo explica)
    return new Promise(res => {
      $('#mbT').textContent = titulo; $('#mbB').innerHTML = html; $('#mbE').textContent = '';
      $('#mbBt').innerHTML = botones.map(b => `<button class="vb ${b[2] ? 'def' : ''}" data-b="${b[0]}">${b[1]}</button>`).join('');
      $('#ov').classList.add('on');
      const h = e => {
        const b = e.target.closest && e.target.closest('[data-b]'); if (!b) return;
        const id = b.dataset.b; const err = validar ? validar(id) : null;
        if (err) { $('#mbE').textContent = err; return; }
        cerrar(id);
      };
      const cerrar = id => { $('#mbBt').removeEventListener('click', h); $('#ov').classList.remove('on'); dlgCierra = null; res(id); };
      dlgCierra = cerrar;
      $('#mbBt').addEventListener('click', h);
      const f = $('#mbB').querySelector('input,select,textarea'); if (f) f.focus();
    });
  }
  const aviso = (m, t) => dialogo(t || 'PNF', '<div>' + m + '</div>', [['ok', 'Aceptar', true]]);
  const valor = id => { const e = $(id); return e ? e.value.trim() : ''; };

  /* ── pantallas ── */
  const frame = (cap, body, st) => `<fieldset class="fr" style="${st || ''}"><legend>${cap}</legend>${body}</fieldset>`;
  const btn = (t, attr, cls) => `<button class="vb ${cls || ''}" ${attr || ''}>${t}</button>`;
  const dl = (k, v) => `<dt>${k}</dt><dd>${v}</dd>`;
  const ETAPAS = {};

  const menu = () => (esTit() ? [] : [['sup', 'Supervisión']]).concat([['oficios', 'Oficios'], ['lectura', 'Lectura del oficio'], ['analisis', 'Análisis técnico'], ['memo', 'Memo'], ['sgd', 'Respuesta en SGD'], '|', ['excel', 'Excel acumulado'], ['config', 'Configuración']]);

  function vFutura(v) {
    const lab = menu().find(m => m[0] === v)[1];
    return frame(lab, `<div class="vacio"><b>Esta pantalla se activa en la etapa ${ETAPAS[v]}.</b><br>Ya está diseñada en el bosquejo; se construye en esa etapa.</div>`);
  }
  const gate = () => `<div class="pinbox"><div style="font-size:15px;font-weight:bold">🔒 Sección protegida</div><div style="margin:6px 0">Solo el administrador conoce el PIN.<br>Titular actual: <b>${esc(R.titular)}</b> (desde ${esc(R.desde)}).</div><div class="row" style="justify-content:center"><input type="password" id="pinIn" inputmode="numeric" maxlength="8" placeholder="PIN" style="width:110px;text-align:center">${btn('Entrar', 'data-a="pin"', 'def')}</div><div id="pinMsg" style="color:#B00020;min-height:18px;margin-top:6px"></div><div style="font-size:11px;color:#555">Se vuelve a bloquear sola a los 10 minutos o con «Bloquear ahora».</div></div>`;

  function vConfig() {
    const t = esTit(); if (!t && !['lec', 'resp', 'act', 'usr'].includes(S.cfgTab)) S.cfgTab = 'resp';
    const c = S.cfgTab, tb = (k, n) => `<span class="mtab ${c === k ? 'on' : ''}" data-ct="${k}">${n}</span>`, bloq = !pinVigente();
    const tabs = (t ? tb('cat', 'Análisis técnico') + tb('memo', 'Memo') + tb('sgd', 'SGD') : '') + tb('lec', (bloq ? '🔒 ' : '') + 'Lectura y expediente') + tb('resp', 'Respaldo') + tb('act', 'Actualizaciones') + tb('usr', (bloq ? '🔒 ' : '') + 'Usuarios y roles');
    let cu = '';
    if (c === 'cat') cu = cfgCat(); else if (c === 'memo') cu = cfgMemo(); else if (c === 'sgd') cu = cfgSgd(); else if (c === 'lec') cu = cfgLec(); else if (c === 'resp') cu = cfgResp(); else if (c === 'act') cu = cfgAct(); else cu = cfgUsr();
    if ((c === 'lec' || c === 'usr')) cu = bloq ? gate() : `<div class="row" style="margin-bottom:6px"><span class="rolchip">🔓 Sección desbloqueada</span>${btn('Bloquear ahora', 'data-a="bloq"', 'peq')}</div>` + cu;
    return `<div class="mtabs" style="margin-bottom:-1px">${tabs}</div><div class="mpage" style="min-height:0;padding:10px">${cu}</div>`;
  }
  const reglaDe = x => x.porDefecto ? 'inf' : x.dejar ? 'dejar' : 'sug';
  function cfgCat() {
    return `<div class="note">Cada fila es un análisis posible. Las <b>frases clave</b> (separadas por coma) hacen que se sugiera solo cuando aparecen en el texto del oficio. Se evalúan de arriba hacia abajo. La columna <b>Regla</b> dice qué hace cada fila (lo explica debajo de cada una). Se puede agregar, cambiar, ordenar o quitar cualquiera.</div>
    <div class="lv" style="max-height:340px"><table><tr><th>Texto del análisis (sale en el memo)</th><th>Frases clave</th><th>Regla</th><th>Orden</th></tr>${D.catalogo.map((x, i) => `<tr><td style="width:46%"><textarea data-cat-t="${i}">${esc(x.t)}</textarea>${x.dejar ? '<div style="font-size:11px">sin texto: la fila queda «Sin asignar» para que la decidas tú</div>' : ''}</td><td><input type="text" data-cat-k="${i}" value="${esc(x.kw.join(', '))}" style="width:100%"></td><td style="min-width:190px"><select data-cat-r="${i}">${[['sug', 'Sugerir'], ['dejar', 'Dejar sin asignar'], ['inf', 'Por defecto (infracciones)']].map(o => `<option value="${o[0]}" ${reglaDe(x) === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select><div style="font-size:11px;margin-top:2px">${{ sug: 'Se sugiere cuando el oficio dice alguna frase clave.', dejar: 'Si aparece una frase clave NO se sugiere nada: queda «Sin asignar» para que decidas tú.', inf: 'Se sugiere si no coincide ninguna frase y el bloque es una INFRACCIÓN. Solo puede haber una fila así.' }[reglaDe(x)]}</div></td><td style="white-space:nowrap">${btn('↑', 'data-a="cat-up" data-i="' + i + '"', 'peq')} ${btn('✕', 'data-a="cat-del" data-i="' + i + '"', 'peq')}</td></tr>`).join('')}</table></div>
    <div class="row" style="margin-top:8px">${btn('+ Agregar análisis', 'data-a="cat-add"', 'def')}<span style="font-size:11.5px">Los cambios se guardan solos.</span></div>`;
  }
  const upInp = (k, v) => `<input type="text" data-cf="${k}" data-up="1" value="${esc(v)}" style="width:100%;text-transform:uppercase">`;
  function cfgMemo() {
    const C = D.cfg;
    return `<div class="g2"><div>${frame('Encabezado del memo', `<label class="lb">DE · nombre</label>${upInp('de.nombre', C.de.nombre)}<label class="lb">DE · cargo</label>${upInp('de.cargo', C.de.cargo)}<label class="lb">PARA · nombre</label>${upInp('para.nombre', C.para.nombre)}<label class="lb">PARA · cargo</label>${upInp('para.cargo', C.para.cargo)}<div style="font-size:11.5px;margin-top:6px">Los datos del encabezado se escriben siempre en <b>MAYÚSCULAS</b>, como en el memo oficial (PNF las convierte al escribir). Para una subrogancia, el PARA también se podrá cambiar solo para un memo desde la pestaña Memo.</div>`)}</div><div>` +
      frame('Textos', `<label class="lb">Asunto (puede usar {{TIPO}}, {{NUMERO}} y {{DOC}})</label><input type="text" data-cf="asunto" value="${esc(C.asunto)}" style="width:100%"><label class="lb">Párrafo inicial</label><textarea data-cf="intro">${esc(C.intro)}</textarea><label class="lb">Cierre</label><input type="text" data-cf="cierre" value="${esc(C.cierre)}" style="width:100%"><div style="font-size:11.5px;margin-top:6px">Sin número ni fecha: los pone el SGD.<br><b>{{TIPO}}</b> = Carta, Memorándum, Oficio… tal como lo dice el oficio · <b>{{NUMERO}}</b> = N° del oficio · <b>{{DOC}}</b> = «la Carta N° 6058/2026 PNF» (con su artículo). PNF los toma de cada oficio al leerlo; si no reconoce el tipo, lo avisa y te pide elegirlo.</div>`) + '</div></div>' +
      frame('Formato de la hoja (como el memo oficial)', `<div class="row"><img src="${LOGO}" alt="SEREMI" style="height:54px"><div style="font-size:11.5px"><b>Logo:</b> se repite en cada página · <b>Hoja:</b> oficio (8,5 × 14 in) · <b>Letra:</b> Verdana (8 pt cabecera y texto, 7 pt tablas) · <b>Tablas:</b> 7 columnas de ancho fijo, encabezado gris · <b>Columna central:</b> el encabezado se copia tal cual del oficio (INFRACCIÓN, DETALLE DE LA FISCALIZACIÓN u OBSERVACIONES)</div></div>`);
  }
  const comboP = (id, val) => `<select data-cf="${id}">${D.cfg.personas.map(p => `<option value="${esc(p.usuario)}" ${p.usuario === val ? 'selected' : ''}>${esc(p.nombre + ' (' + p.usuario + ')')}</option>`).join('')}</select>`;
  function cfgSgd() {
    const C = D.cfg, F = C.form;
    return `<div class="g2"><div>${frame('Personas del SGD', `<div style="font-size:11.5px;margin-bottom:6px">El SGD busca por usuario; se guarda también el nombre completo para elegir la fila correcta (nunca por posición).</div><div class="lv" style="height:auto;max-height:170px"><table><tr><th>Usuario</th><th>Nombre completo</th><th>Cargo</th><th></th></tr>${C.personas.map((p, i) => `<tr><td class="mono">${esc(p.usuario)}</td><td>${esc(p.nombre)}</td><td>${esc(p.cargo)}</td><td>${btn('Quitar', 'data-a="per-del" data-i="' + i + '"', 'peq')}</td></tr>`).join('')}</table></div><div class="row" style="margin-top:8px"><input type="text" id="nuPU" placeholder="usuario" style="width:90px"><input type="text" id="nuPN" placeholder="NOMBRE COMPLETO" style="width:190px;text-transform:uppercase"><input type="text" id="nuPC" placeholder="CARGO" style="width:140px;text-transform:uppercase">${btn('Agregar', 'data-a="per-add"', 'def peq')}</div><div style="font-size:11.5px;margin-top:6px">En la etapa 5 el nombre y cargo se traerán del SGD al verificar.</div>`)}
      ${frame('Quién visa y quién firma', `<label class="lb">Visador (siempre el mismo)</label>${comboP('visador', C.visador)}<label class="lb">Firmante (siempre el mismo)</label>${comboP('firmante', C.firmante)}`)}</div><div>
      ${frame('Valores fijos del formulario del SGD', `<label class="lb">Tipo de respuesta</label><input type="text" data-cf="form.tipoResp" value="${esc(F.tipoResp)}" style="width:100%"><label class="lb">Tipo de documento</label><input type="text" data-cf="form.tipoDoc" value="${esc(F.tipoDoc)}" style="width:100%"><label class="lb">Tipo de materia</label><input type="text" data-cf="form.tipoMateria" value="${esc(F.tipoMateria)}" style="width:100%"><label class="lb">Ciudad</label><input type="text" data-cf="form.ciudad" value="${esc(F.ciudad)}" style="width:100%"><label class="lb">Actos con efecto sobre terceros</label><input type="text" data-cf="form.actos" value="${esc(F.actos)}" style="width:100%"><div style="font-size:11.5px;margin-top:6px">La <b>Materia</b> no va aquí: es el ANT. de cada oficio.</div>`)}
      ${frame('Distribución', `<label class="lb">Lista guardada en el SGD («Mis Listas»)</label><input type="text" data-cf="listaDist" value="${esc(C.listaDist)}" style="width:100%"><div style="font-size:11.5px;margin:6px 0">Créala una vez en el SGD. PNF la busca por nombre: si no existe o está vacía, se detiene y avisa.</div>`)}
      ${frame('Al terminar', `<label><input type="checkbox" data-cf="pausaGuardar" ${C.pausaGuardar ? 'checked' : ''}> Pausar antes de «Guardar» para revisar el formulario</label><div style="font-size:11.5px;margin-top:6px">«Despachar» siempre es manual: PNF baja hasta el botón y lo resalta.</div>`)}</div></div>`;
  }
  function cfgLec() {
    return (esTit() ? '' : `<div class="note inf">ℹ La clave y el motor de lectura solo se usan en el computador del <b>titular</b> (${esc(R.titular)}), que es quien lee los oficios. Aquí quedan guardados en <b>este</b> computador por si algún día eres titular.</div>`) + `<div class="g2">${frame('Lectura del PDF', `<label class="lb">Motor</label><select data-lc="motor"><option value="gemini" ${L.motor === 'gemini' ? 'selected' : ''}>Gemini (con lector normal de respaldo)</option><option value="pdfjs" ${L.motor === 'pdfjs' ? 'selected' : ''}>Solo lector normal (pdf.js)</option></select><label class="lb">Modelo</label><input type="text" data-lc="modelo" value="${esc(L.modelo)}" style="width:100%"><label class="lb">Clave de Gemini</label><input type="password" data-lc="clave" value="${esc(L.clave)}" style="width:100%" autocomplete="off"><div style="font-size:11.5px;margin-top:6px">La clave se guarda en Tampermonkey, solo en este computador: no va en el código de GitHub ni en el respaldo. Los oficios se envían a Google para leerlos.</div><div class="row" style="margin-top:6px">${btn('Probar conexión', 'data-a="gem-test"', 'peq')}<span id="gemRes" style="font-size:11.5px"></span></div>`)}
      ${frame('Expediente válido', `<div class="mono" style="padding:6px;background:#fff;border:1px solid var(--ib)">E + número + / + año de 4 dígitos<br>ej. E155797/2026 · se busca en la última página</div><div style="font-size:11.5px;margin-top:6px">Regla estricta: cualquier otra forma no se acepta. Nunca lo decide la IA.</div>`)}</div>
      ${frame('Cambiar la clave de Gemini (paso a paso)', '<ol style="margin:0;padding-left:20px;line-height:1.6"><li>Entra a <b>aistudio.google.com/apikey</b> con la cuenta de Google que quieras usar.</li><li>Pulsa <b>Create API key</b> (elige o crea un proyecto) y copia la clave.</li><li>Pégala arriba en <b>Clave de Gemini</b> y pulsa <b>Probar conexión</b>.</li><li>Si la prueba resulta bien, queda guardada en Tampermonkey de este computador y reemplaza a la anterior. No hay que editar el script, así la clave nunca se sube a GitHub.</li></ol><div style="font-size:11.5px;margin-top:6px">El cupo de uso pertenece al proyecto de esa clave: si es la misma que usa Cierres Waze, comparten el límite.</div>')}`;
  }
  function cfgResp() {
    const t = esTit(), carp = { ok: '<span class="estado e-v">conectada</span>', sin: '<span class="estado e-a">sin conectar</span>', permiso: '<span class="estado e-a">falta confirmar el permiso</span>' }[S.carpeta];
    return `<div class="g2"><div>${frame('Carpeta de respaldo y de datos compartidos', `<div class="note">${t ? 'PNF reescribe sola los archivos cada vez que se guarda algo importante. El analista no tiene que hacer nada.' : 'PNF lee de esta carpeta los datos de ' + esc(R.titular) + '. No escribe ahí los datos del titular.'}</div><dl class="kv"><dt>Carpeta</dt><dd>${carp} <span class="mono">${esc(S.carpetaNombre)}</span></dd>${t ? dl('Archivos', '<span class="mono">PNF-datos.json</span> (datos) · <span class="mono">PNF-version.json</span> (aviso de cambios, diminuto) · <span class="mono">copias/</span> (últimas 10)') : ''}${dl('Último respaldo', cuando(L.ultResp))}${dl('No incluye', 'la clave de Gemini: se guarda solo en cada computador. ' + btn('Ir a la clave de Gemini…', 'data-ct="lec"', 'peq'))}</dl><div class="row" style="margin-top:8px">${btn(S.carpeta === 'sin' ? 'Elegir carpeta…' : 'Cambiar carpeta…', 'data-a="carpeta"')}${S.carpeta === 'permiso' ? btn('Permitir', 'data-a="permiso"', 'def') : ''}${t ? btn('Respaldar ahora', 'data-a="resp-ahora"', 'def') + btn('Restaurar desde la carpeta…', 'data-a="adoptar"') : ''}</div>`)}</div><div>
      ${frame('Que el respaldo llegue a Google Drive (una sola vez)', '<ol style="margin:0;padding-left:20px;line-height:1.6"><li>Instala <b>Google Drive para escritorio</b> e inicia sesión con tu cuenta.</li><li>En Drive crea una carpeta, por ejemplo <b>Mi unidad\\Unidad Analisis\\PNF</b>.</li><li>Aquí pulsa <b>Elegir carpeta…</b> (o «Cambiar carpeta…» si ya hay una) y elige esa carpeta. Desde ese momento PNF escribe ahí sin pedir nada.</li><li>Si Chrome vuelve a pedir permiso al reabrir el navegador, aparece un aviso con el botón <b>Permitir</b>. Mientras no se confirme, PNF guarda una copia en <b>Descargas\\PNF-respaldos</b> para no quedar sin respaldo.</li></ol>')}
      ${frame('Usar PNF en otro computador', '<ol style="margin:0;padding-left:20px;line-height:1.6"><li>Instala <b>Tampermonkey</b> y abre el enlace de PNF en GitHub para instalarlo.</li><li>Instala <b>Google Drive para escritorio</b> con la misma cuenta, para que aparezca la carpeta PNF.</li><li>Abre el SGD y pulsa <b>Elegir carpeta…</b> señalando la carpeta de Drive.</li><li>PNF encuentra los datos y, si eres el titular, ofrece restaurarlos. Después pega la clave de Gemini.</li></ol>')}</div></div>`;
  }
  function cfgAct() {
    const info = (typeof GM_info !== 'undefined' && GM_info.script) ? GM_info.script.version : VERSION;
    return `<div class="g2"><div>${frame('Versión', `<dl class="kv">${dl('Instalada', '<b>v' + esc(info) + '</b>')}${dl('Última en GitHub', '<span id="actUlt">' + (S.actTxt || 'sin comprobar') + '</span>')}</dl><div class="row" style="margin-top:8px">${btn('Buscar actualizaciones', 'data-a="act-buscar"', 'def')}</div>`)}</div><div>${frame('Cómo se actualiza PNF', '<ol style="margin:0;padding-left:20px;line-height:1.6"><li>El código de PNF vive en un repositorio público de GitHub (sin datos ni claves).</li><li>El encabezado del script lleva <b>@updateURL</b> y <b>@downloadURL</b> apuntando a ese archivo.</li><li>Tampermonkey revisa solo cada cierto tiempo (se puede cambiar en sus Ajustes) y, si el <b>@version</b> es mayor, instala la nueva versión.</li><li>Tus datos no se tocan: están guardados aparte del código.</li></ol><div style="font-size:11.5px;margin-top:6px">Primera instalación: abrir una vez el enlace del archivo en GitHub y pulsar «Instalar».</div>')}</div></div>`;
  }
  function cfgUsr() {
    const adm = esAdm(), us = Object.keys(R.roles).filter(u => !S.uFil || u.includes(S.uFil.toLowerCase()));
    const pasar = R.titular !== R.base;
    return `<div class="g2"><div>${frame('Titular actual', `<div class="row"><span class="rolchip">Titular</span><b style="font-size:14px">${esc(R.titular)}</b></div><dl class="kv" style="margin-top:8px">${dl('Desde', esc(R.desde))}${dl('Asignado por', esc(R.por))}${dl('Hasta', R.hasta ? esc(R.hasta) : 'sin fecha (manual)')}</dl><div style="font-size:11.5px;margin-top:6px">Solo el titular trabaja los datos y envía al SGD. Los demás miran, dejan observaciones y agregan oficios.</div>`)}
      ${frame('Traspasar el trabajo', adm ? `<div style="font-size:11.5px;margin-bottom:6px">Para vacaciones, licencia o una persona nueva. Siempre hay una sola persona titular.</div><div class="row">${btn('Traspasar el trabajo…', 'data-a="tras"', 'def')}${pasar ? btn('Devolver a ' + esc(R.base), 'data-a="tras-dev"') : ''}</div>` : `<div style="font-size:11.5px">Solo el administrador (${esc(R.admins.join(', '))}) puede traspasar la titularidad.</div>`)}
      ${frame('Alertas', `<dl class="kv" style="grid-template-columns:190px minmax(0,1fr)"><dt>Expediente sin responder desde</dt><dd><input type="number" min="1" data-ua="dias" value="${R.dias}" style="width:60px" ${adm ? '' : 'disabled'}> días</dd><dt>Datos antiguos desde</dt><dd><input type="number" min="1" data-ua="horas" value="${R.horas}" style="width:60px" ${adm ? '' : 'disabled'}> horas sin guardar</dd><dt>Buscar avisos nuevos cada</dt><dd><input type="number" min="10" data-ua="seg" value="${R.seg}" style="width:60px" ${adm ? '' : 'disabled'}> segundos</dd></dl><div style="font-size:11.5px;margin-top:6px">Solo lee el archivo diminuto de versión (ocupa casi nada) y se pausa si la pestaña está oculta. ${adm ? 'Los cambios los escribe el administrador en PNF-roles.json.' : 'Solo el administrador cambia estos valores.'}</div>`)}</div><div>
      ${frame('Usuarios y roles (' + Object.keys(R.roles).length + ')', (Object.keys(R.roles).length > 6 ? `<input type="text" data-ufil="1" value="${esc(S.uFil)}" placeholder="Buscar usuario…" style="width:100%;margin-bottom:6px">` : '') + `<div class="lv" style="height:auto;max-height:200px"><table><tr><th>Usuario en PNF</th><th>Nombre como lo muestra el SGD</th><th>Rol</th>${adm ? '<th></th>' : ''}</tr>${us.map(u => `<tr><td class="mono">${esc(u)}${u === R.titular ? ' <span class="rolchip">titular</span>' : ''}</td><td>${adm ? `<input type="text" data-nom="${esc(u)}" value="${esc(R.nombres[u] || '')}" placeholder="NOMBRE APELLIDO" style="width:100%;text-transform:uppercase">` : esc(R.nombres[u] || '—')}</td><td>${adm ? `<select data-rol="${esc(u)}">${ROLES_TXT.map(r => `<option ${R.roles[u] === r ? 'selected' : ''}>${r}</option>`).join('')}</select>` : esc(R.roles[u])}</td>${adm ? '<td>' + (u === R.titular ? '' : btn('Quitar', 'data-a="usr-del" data-u="' + esc(u) + '"', 'peq')) + '</td>' : ''}</tr>`).join('')}</table></div>` + (adm ? `<div class="row" style="margin-top:8px"><b>Agregar usuario</b><input type="text" id="nuUsr" placeholder="usuario (ej. cnunez)" style="width:130px"><input type="text" id="nuNom" placeholder="NOMBRE como lo muestra el SGD" style="width:220px;text-transform:uppercase"><select id="nuRol"><option>Supervisor</option><option>Analista</option><option>Administrador</option><option>Espectador</option></select>${btn('Agregar', 'data-a="usr-add"', 'def')}</div>` : '') + '<div style="font-size:11.5px;margin-top:6px">PNF reconoce a la persona por el <b>nombre con el que entró al SGD</b> («Usuario conectado: …», arriba a la derecha), sin claves propias. El SGD no muestra el usuario ni el apellido materno: por eso se anota el nombre tal como aparece. Los roles son una convención de confianza del equipo, no una seguridad real.</div>')}
      ${frame('Historial de traspasos' + (adm ? ' (' + R.hist.length + ')' : ''), adm ? `<div class="lv" style="height:auto;max-height:150px"><table><tr><th>Cuándo</th><th>Titular</th><th>Por</th><th>Motivo</th><th>Hasta</th></tr>${R.hist.slice(0, 50).map(h => `<tr><td>${esc(h.f)}</td><td>${esc(h.a)}</td><td>${esc(h.por)}</td><td>${esc(h.motivo)}</td><td>${esc(h.hasta)}</td></tr>`).join('')}</table></div><div style="font-size:11.5px;margin-top:4px">Se muestran los últimos 50; el archivo guarda todos.</div>` : '<div style="font-size:11.5px">Solo el administrador ve el historial.</div>')}
      ${frame('Quién eres en PNF', `<dl class="kv">${dl('Nombre en el SGD', S.sgd ? '<b>' + esc(S.sgd.nombre) + '</b> <small>(' + esc(S.sgd.cargo) + ')</small>' : 'no se ve en esta página')}${dl('Coincide con', S.sgd ? (candidatosPorNombre(S.sgd.nombre).map(u => '<span class="mono">' + esc(u) + '</span>').join(', ') || 'nadie de la lista') : '—')}${dl('Tú en PNF', esc(L.usuario || '—'))}${dl('Página', esc(document.title.slice(0, 60)))}</dl><div class="row" style="margin-top:6px">${btn('Volver a leer', 'data-a="detectar"', 'peq')}${btn('Cambiar quién soy…', 'data-a="ident"', 'peq')}</div>`)}</div></div>`;
  }

  function render() {
    const adm = esAdm(), t = esTit();
    $('#toolbar').innerHTML = (t ? '<button class="tb" data-a="guardar">💾 Guardar</button><span class="tsep"></span>' : aporta() ? '<button class="tb" data-a="examinar">➕ Agregar oficio</button><span class="tsep"></span>' : '') + (t ? '<button class="tb" data-a="verif">🔄 Verificar bandeja</button>' : '') + '<button class="tb" data-a="act">↻ Actualizar datos</button><span class="tsep"></span><button class="tb" data-go="excel">📊 Excel acumulado</button><span class="tsep"></span><button class="tb" data-go="config">⚙ Configuración</button>';
    $('#tabs').innerHTML = menu().map(m => m === '|' ? '<span class="sepx"></span>' : `<button class="mtab ${S.vista === m[0] ? 'on' : ''} ${ETAPAS[m[0]] ? 'fut' : ''}" data-go="${m[0]}">${m[1]}</button>`).join('');
    const al = alertas().map(a => `<div class="note ${a[0]}">${a[1]}</div>`).join('');
    let h;
    if (S.vista === 'oficios') h = al + frameAportes() + (!t && L.usuario ? `<div class="ro">👁 Estás viendo los datos de <b>${esc(R.titular)}</b> en solo lectura.</div>` : '') + vOficios();
    else if (S.vista === 'sup') h = al + vSup();
    else if (S.vista === 'lectura') h = al + roBanner() + vLectura();
    else if (S.vista === 'analisis') h = al + vAnalisis();
    else if (S.vista === 'memo') h = al + vMemo();
    else if (S.vista === 'sgd') h = al + vSgd();
    else if (S.vista === 'excel') h = vExcel();
    else if (S.vista === 'config') h = vConfig();
    else h = vFutura(S.vista);
    const sc = $('.form').scrollTop, scs = [...root.querySelectorAll('#vista .lv, #vista .tw')].map(x => [x.scrollTop, x.scrollLeft]);   // las tablas con barra propia también conservan su posición
    $('#vista').innerHTML = h; $('.form').scrollTop = sc; if (S.vista === 'oficios') renderTabla();
    root.querySelectorAll('#vista .lv, #vista .tw').forEach((x, k) => { if (scs[k]) { x.scrollTop = scs[k][0]; x.scrollLeft = scs[k][1]; } });
    if (S.foco) { const f = S.foco; S.foco = null; const tr = $('#vista tr[data-row="' + f.i + '"]'); if (tr) { tr.scrollIntoView({ block: 'nearest' }); const el = tr.querySelector(f.que === 'sel' ? '[data-catb]' : f.que === 'cus' ? '[data-cus]' : '[data-a="conf"]') || tr.querySelector('[data-catb]'); if (el) el.focus({ preventScroll: true }); tr.classList.add('foco'); } }
    pintarEstado();
    // el ícono cuenta solo expedientes sin responder (incluye los devueltos): rojo = vencidos (días de la configuración o más), verde = dentro del plazo
    const cp = conteoPlazo(), bv = $('#fabN'), bp = $('#fabP'); bv.textContent = cp.v; bv.classList.toggle('on', cp.v > 0); bp.textContent = cp.p; bp.classList.toggle('on', cp.p > 0);
    const nd = devueltos().length; $('#fabD').classList.toggle('on', nd > 0);   // ⚠ = hay oficios que volvieron a la bandeja
    $('#fab').title = 'PNF · ' + cp.v + ' vencido(s) · ' + cp.p + ' dentro del plazo' + (nd ? ' · ⚠ ' + nd + ' devuelto(s)' : '') + ' · arrastra para mover';
  }
  const claseEstado = s => s === 'Pendiente' ? 'e-z' : s === 'Devuelto' ? 'e-a' : s === 'Por despachar' ? 'e-p' : 'e-v';
  function conteoPlazo() { const os = datos().oficios || {}; let v = 0, p = 0; Object.keys(os).forEach(e => { if (os[e].sgd === 'Respondido') return; if (diasDe(os[e]) >= R.dias) v++; else p++; }); return { v, p }; }
  function pintarEstado() {
    $('#st1').textContent = S.msg;
    $('#st2').innerHTML = 'Usuario: <b>' + esc(L.usuario || '—') + '</b> · ' + etiquetaRol();
    const rev = esTit() ? D.rev : (V.rev || 0), ts = esTit() ? D.guardado : V.guardado;
    $('#st3').innerHTML = 'Datos de ' + esc(R.titular) + ' · revisión ' + rev + ' · ' + cuando(ts) + '<button class="stbtn" data-a="act" title="Volver a leer ahora los datos de la carpeta">↻ Actualizar</button>';
    const g = $('#stG'); if (!esTit()) { g.className = ''; g.textContent = ''; } else { g.className = S.dirty ? 'dirty' : ''; g.textContent = S.dirty ? '● Cambios sin guardar' : 'Guardado ✓ ' + (D.guardado ? cuando(D.guardado) : 'sin cambios'); }
    $('#stV').textContent = 'v' + VERSION;
  }
  function minimizar() { $('#win').classList.remove('on'); $('#fab').classList.add('on'); L.min = true; saveL(); }
  function restaurar() { $('#fab').classList.remove('on'); $('#win').classList.add('on'); L.min = false; saveL(); render(); }

  /* ── acciones ── */
  async function actualizar() {
    toast('Leyendo la versión de los datos en la carpeta…');
    await sincronizar();
    if (esTit() && S.conflicto) toast('⚠ Hay datos más nuevos de otra sesión: usa «Actualizar y revisar».');
    else toast('Datos al día: revisión ' + (esTit() ? D.rev : (V.rev || 0)));
    render();
  }
  async function identificar() {
    S.sgd = leerSesionSGD() || S.sgd; const nom = S.sgd ? S.sgd.nombre : '';
    const sug = L.usuario || (nom ? (candidatosPorNombre(nom)[0] || '') : '');
    const html = `<div>${nom ? 'Entraste al SGD como <b>' + esc(nom) + '</b>' + (S.sgd.cargo ? ' (' + esc(S.sgd.cargo) + ')' : '') + '.' : 'No pude leer tu nombre en esta página del SGD.'}<br>PNF necesita saber <b>quién eres en PNF</b> para darte el rol correcto.</div><label class="lb">Soy</label><select id="idU" style="width:100%"><option value="">— elegir —</option>${Object.keys(R.roles).map(u => `<option value="${esc(u)}" ${u === sug ? 'selected' : ''}>${esc(u)} · ${esc(R.roles[u])}${R.nombres[u] ? ' · ' + esc(R.nombres[u]) : ''}</option>`).join('')}</select><div style="font-size:11.5px;margin-top:6px">Si no apareces en la lista, pídele al administrador que te agregue en Configuración › Usuarios y roles.</div>`;
    const r = await dialogo('¿Quién eres en PNF?', html, [['ok', 'Guardar', true], ['no', 'Cancelar']], id => id === 'ok' && !valor('#idU') ? 'Elige tu usuario de la lista.' : null);
    if (r !== 'ok') return;
    const u = valor('#idU');
    if (nom && R.nombres[u] && !mismoNombre(R.nombres[u], nom)) {
      const c = await dialogo('El nombre no coincide', `<div>En PNF, <b>${esc(u)}</b> figura como <b>${esc(R.nombres[u])}</b>, pero entraste al SGD como <b>${esc(nom)}</b>.<br>¿Seguro que eres ${esc(u)}?</div>`, [['si', 'Sí, continuar'], ['no', 'Volver', true]]);
      if (c !== 'si') return identificar();
    }
    L.usuario = u; L.nombreSGD = nom; saveL(); toast('Eres ' + L.usuario + ' (' + etiquetaRol() + ')'); await sincronizar(); render();
  }
  async function guiaCarpeta() {
    const html = `<div><b>Dónde se guardan los datos</b></div><ol style="margin:6px 0 0;padding-left:20px;line-height:1.6"><li>Crea (una sola vez) una carpeta dentro de tu Google Drive para escritorio, por ejemplo <b>Mi unidad\\Unidad Analisis\\PNF</b>. Si trabajan varias personas, el administrador la comparte con ellas.</li><li>Pulsa <b>Elegir carpeta…</b> y selecciónala en la ventana de Windows.</li><li>Chrome pregunta si permites que este sitio edite archivos de esa carpeta: pulsa <b>Editar archivos</b>.</li></ol><div style="font-size:11.5px;margin-top:8px">Todos eligen <b>la misma carpeta</b>. Chrome no deja elegir Escritorio, Documentos ni Descargas completos (sí una subcarpeta). Se puede cambiar después en Configuración › Respaldo.</div>`;
    const r = await dialogo('Conectar la carpeta de respaldo', html, [['elegir', 'Elegir carpeta…', true], ['luego', 'Más tarde']]);
    if (r === 'elegir') await elegirCarpeta();
  }
  async function adoptar() {
    if (S.carpeta !== 'ok') return aviso('Primero conecta la carpeta.');
    const v = await leerJSON('PNF-version.json'); if (!v) return aviso('En la carpeta todavía no hay datos guardados.');
    const r = await dialogo('Restaurar datos', `<div>En la carpeta hay datos de la <b>revisión ${v.rev}</b> (${cuando(v.ts)}, escritos por ${esc(v.titular || '—')}).<br>Los datos actuales de este equipo (revisión ${D.rev}) se guardan antes como copia de seguridad.</div>`, [['ok', 'Restaurar', true], ['no', 'Cancelar']]);
    if (r !== 'ok') return;
    if (await cargarDeCarpeta('restaurado')) { render(); }
  }
  async function traspasar() {
    const cand = Object.keys(R.roles).filter(u => u !== R.titular && R.roles[u] !== 'Espectador');
    if (!cand.length) return aviso('No hay otros usuarios con rol para traspasar. Agrega uno primero.');
    const html = `<label class="lb">Nuevo titular</label><select id="trQ" style="width:100%">${cand.map(u => '<option>' + esc(u) + '</option>').join('')}</select><label class="lb">Motivo</label><input type="text" id="trM" style="width:100%" placeholder="Ej.: licencia médica, vacaciones, persona nueva"><label class="lb">Hasta</label><label><input type="radio" name="trH" value="m" checked> Manual (hasta que el administrador lo devuelva)</label><br><label><input type="radio" name="trH" value="f"> Hasta la fecha</label> <input type="text" id="trF" placeholder="dd/mm/aaaa" style="width:110px"><div style="font-size:11.5px;margin-top:8px">Antes de traspasar PNF guarda una copia de seguridad. <b>${esc(R.titular)}</b> pasa a solo lectura la próxima vez que abra PNF.</div>`;
    const r = await dialogo('Traspasar el trabajo', html, [['ok', 'Traspasar', true], ['no', 'Cancelar']], id => { if (id !== 'ok') return null; if (!valor('#trM')) return 'Escribe el motivo del traspaso.'; if ($('input[name=trH]:checked').value === 'f' && !/^\d{2}\/\d{2}\/\d{4}$/.test(valor('#trF'))) return 'Escribe la fecha como dd/mm/aaaa.'; return null; });
    if (r !== 'ok') return;
    const q = valor('#trQ'), f = $('input[name=trH]:checked') ? $('input[name=trH]:checked').value : 'm', hasta = f === 'f' ? valor('#trF') : 'manual', mot = valor('#trM');
    await aplicarTraspaso(q, mot, hasta);
    if (q === D.cfg.visador || q === D.cfg.firmante) aviso('<b>' + esc(q) + '</b> figura como ' + (q === D.cfg.visador ? 'visador' : 'firmante') + ' en la configuración del SGD. Revisa Configuración › SGD antes de enviar.');
  }
  async function aplicarTraspaso(q, mot, hasta) {
    try { if (esTit() && S.carpeta === 'ok') await copiaSeguridad(); } catch (e) { /* sigue */ }
    R.hist.unshift({ f: hoyTxt() + ' ' + hh(), a: q, por: L.usuario, motivo: mot, hasta }); R.titular = q; R.desde = hoyTxt(); R.por = L.usuario; R.hasta = hasta === 'manual' ? '' : hasta;
    await guardarRoles('traspaso a ' + q);
  }
  async function probarGemini() {
    const out = $('#gemRes'); if (!L.clave) { out.textContent = 'Escribe primero la clave.'; return; }
    out.textContent = 'Probando…';
    try {
      const r = await http({ method: 'POST', url: 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(L.modelo) + ':generateContent?key=' + encodeURIComponent(L.clave), headers: { 'Content-Type': 'application/json' }, data: jsonStr({ contents: [{ role: 'user', parts: [{ text: 'Responde solo: OK' }] }] }) });
      if (r.status === 200) out.innerHTML = '<span class="estado e-v">✔ conexión correcta</span> ' + esc(L.modelo);
      else { let m = ''; try { m = JSON.parse(r.text).error.message; } catch (e) { /* sin detalle */ } out.innerHTML = '<span class="estado e-r">✖ error ' + r.status + '</span> ' + esc(m.slice(0, 140)); }
    } catch (e) { out.innerHTML = '<span class="estado e-r">✖ ' + esc(e.message) + '</span>'; }
  }
  async function buscarActualizacion() {
    const ult = $('#actUlt');
    if (/USUARIO\/REPO/.test(URL_ACTUALIZAR)) { S.actTxt = 'falta indicar el repositorio de GitHub (se configura al subir el script)'; if (ult) ult.textContent = S.actTxt; return; }
    if (ult) ult.textContent = 'consultando…';
    try {
      const r = await http({ url: URL_ACTUALIZAR + '?t=' + Date.now() }); const m = r.text.match(/@version\s+([0-9.]+)/);
      if (r.status !== 200 || !m) S.actTxt = 'no se pudo leer (' + r.status + ')';
      else S.actTxt = cmpVer(m[1], VERSION) > 0 ? '<b>v' + m[1] + '</b> disponible — Tampermonkey la instalará sola; o ábrela en GitHub y pulsa «Instalar»' : 'v' + m[1] + ' · estás al día ✓';
    } catch (e) { S.actTxt = 'sin conexión (' + esc(e.message) + ')'; }
    render();
  }
  const guardaCfg = motivo => guardar(motivo || 'configuración');

  root.addEventListener('click', async e => {
    const el = e.target.closest('[data-a],[data-go],[data-ct],[data-cbpick],[data-cbtog],[data-ord],[data-pg],[data-fila]'); if (!el) { if (!e.target.closest('.cb')) { cbCerrar(); $$('.cbIn').forEach(i => { const o = entradaPend(); i.value = i.dataset.cbin === 'selXl' ? S.xlExp : (S.vista === 'lectura' && o ? o.exp : exAct()); }); } return; } const d = el.dataset;
    if (d.cbpick !== undefined) { cbPick(d.cbpick, d.v); return; }
    if (d.cbtog) { const l = $('#cbl-' + d.cbtog); if (l.classList.contains('on')) cbCerrar(); else { cbCerrar(); cbAbrir(d.cbtog, '', false); } return; }
    if (d.ord) { if (S.tb.ord.c === d.ord) S.tb.ord.d = -S.tb.ord.d; else S.tb.ord = { c: d.ord, d: 1 }; renderTabla(); return; }
    if (d.pg) { S.tb.pag = +d.pg; renderTabla(); return; }
    if (d.fila) { S.sel = d.fila; S.exp = d.fila; $$('#tbBody tr.sel').forEach(r => r.classList.remove('sel')); el.classList.add('sel'); return; }   // solo marca la fila: si se redibujara la tabla, el 2.º clic caería en otra fila y no habría doble clic
    if (d.go) { if (S.dirty && d.go !== S.vista && esTit()) await guardar('al cambiar de pantalla'); S.vista = d.go; if (d.go === 'config' && !esTit()) S.cfgTab = 'resp'; render(); return; }
    if (d.ct) { S.cfgTab = d.ct; render(); return; }
    const i = +d.i;
    switch (d.a) {
      case 'guardar': await guardar('manual', true); break;
      case 'act': await actualizar(); break;
      case 'sgd-enviar': await enviarSGD(); break;
      case 'pdf-ver': abrirPdf(); break;
      case 'pdf-cerrar': cerrarPdf(); break;
      case 'ir-despachar': S.exp = d.id; S.sel = d.id; await enviarSGD('ir'); break;
      case 'crear-nuevo': S.exp = d.id; S.sel = d.id; await enviarSGD('nuevo'); break;
      case 'verif': await verificarBandeja(); break;
      case 'atend': if (D.oficios[d.id]) { D.oficios[d.id].atendido = true; await guardar('aviso atendido'); } break;
      case 'run-sigue': if (RUN) { RUN.estado = 'corriendo'; saveRun(); tarjeta(); try { await guardarFormulario(); } catch (x) { /* ya informado */ } } break;
      case 'run-reintenta': if (RUN) { RUN.reintento = true; saveRun(); await correrRun(); } break;
      case 'run-cancela': cancelarRun(); break;
      case 'conf': case 'conf-sug': case 're': case 'bulk': case 'ir-cat': await accionAnalisis(d.a, d); break;
      case 'ap-nuevo': await nuevoAporte(i); break;
      case 'apr': await abrirApv(d.id); break;
      case 'apv': await marcarAporte(d.id, 'visto', 'aporte visto'); render(); break;
      case 'apa': { const a = S.aportes.find(x => x.id === d.id), r = a && D.oficios[a.exp] && D.oficios[a.exp].bloques[a.fila]; if (r) { r.cat = a.cat; r.custom = ''; r.st = 'sug'; } await marcarAporte(d.id, 'aplicada', 'aporte aplicado'); render(); break; }
      case 'apd': await marcarAporte(d.id, 'descartada', 'aporte descartado'); render(); break;
      case 'apc': await marcarAporte(d.id, 'corregido', 'error del memo corregido'); render(); break;
      case 'vis': L.vistos = L.vistos || {}; L.vistos[d.id] = 1; saveL(); render(); break;
      case 'memo-pdf': await descargarMemo('pdf'); break;
      case 'memo-doc': await descargarMemo('doc'); break;
      case 'ir-memo-cfg': S.vista = 'config'; S.cfgTab = 'memo'; render(); break;
      case 'para-memo': await cambiarPara(); break;
      case 'para-quitar': { const o = ofAct(); if (o) { delete o.paraMemo; o.memoOk = false; await guardar('PARA vuelve al de Configuración'); render(); } break; }
      case 'err-memo': await reportarErrorMemo(); break;
      case 'xls': await exportarExcel(); break;
      case 'examinar': $('#fileIn').click(); break;
      case 'lec-ir': S.lec = { pend: d.id }; S.vista = 'lectura'; render(); break;
      case 'lec-agregar': await agregarPend(d.id); break;
      case 'lec-analisis': { const ex = (S.lec && S.lec.exp) || S.sel; if (ex) { S.exp = ex; S.sel = ex; } S.vista = 'analisis'; render(); break; }
      case 'lec-descartar': await descartarPend(d.id); break;
      case 'lec-usa': { const pe = entradaPend(); if (pe) { pe.usa = d.m; render(); } break; }
      case 'lec-ia': { const pe = entradaPend(); if (!pe) break; toast('Leyendo con IA…'); try { pe.ia = await leerConIA(pe.bytes); pe.errIA = ''; pe.modelo = L.modelo; pe.usa = 'ia'; pe.cmp = comparar(pe.ia, pe.normal); toast('Lectura con IA lista'); } catch (x) { await aviso('No se pudo leer con IA: ' + esc(x.message || String(x))); } render(); break; }
      case 'lec-corregir': await corregirDatos(); break;
      case 'ver-pdf': await verPdf(); break;
      case 'abrir-sel': if (!S.sel) { await aviso('Elige primero una fila de la lista.'); break; } S.lec = { exp: S.sel }; S.vista = 'lectura'; render(); break;
      case 'quitar-sel': if (!S.sel) { await aviso('Elige primero una fila de la lista.'); break; } await quitarOficio(S.sel); break;
      case 'visto': L.visto = S.ver ? S.ver.rev : 0; saveL(); render(); break;
      case 'ident': await identificar(); break;
      case 'detectar': { const q = verificarSesion(); if (q === 'cambio') await identificar(); render(); break; }
      case 'carpeta': await elegirCarpeta(); break;
      case 'permiso': if (await conectar(true)) { await sincronizar(); toast('Carpeta reconectada'); } render(); break;
      case 'adoptar': await adoptar(); break;
      case 'resp-ahora': await guardar('respaldo manual', true); break;
      case 'pin': if (pinOk(valor('#pinIn'))) { S.pinOk = true; guardaPIN(); toast('Sección desbloqueada'); render(); } else { $('#pinMsg').textContent = 'PIN incorrecto.'; $('#pinIn').value = ''; $('#pinIn').focus(); } break;
      case 'bloq': S.pinOk = false; toast('Sección bloqueada'); render(); break;
      case 'cat-add': { D.catalogo.push({ id: 'n' + Date.now(), t: 'Nuevo análisis…', kw: [] }); await guardaCfg('catálogo de análisis'); const ta = $('[data-cat-t="' + (D.catalogo.length - 1) + '"]'); if (ta) { const lv = ta.closest('.lv'); if (lv) lv.scrollTop = lv.scrollHeight; ta.focus(); ta.select(); } break; }
      case 'cat-up': if (i > 0) { [D.catalogo[i - 1], D.catalogo[i]] = [D.catalogo[i], D.catalogo[i - 1]]; await guardaCfg('catálogo de análisis'); } break;
      case 'cat-del': if (D.catalogo.length > 1) { D.catalogo.splice(i, 1); await guardaCfg('catálogo de análisis'); } break;
      case 'per-add': { const u = valor('#nuPU').toLowerCase(), n = valor('#nuPN').toUpperCase(), c = valor('#nuPC').toUpperCase(); if (!/^[a-z][a-z0-9._-]{2,}$/.test(u) || !n) { aviso('Escribe el usuario del SGD y el nombre completo.'); break; } if (D.cfg.personas.some(p => p.usuario === u)) { aviso('Ese usuario ya está en la lista.'); break; } D.cfg.personas.push({ usuario: u, nombre: n, cargo: c, ok: false }); await guardaCfg('personas del SGD'); break; }
      case 'per-del': { const p = D.cfg.personas[i]; if (!p) break; if (p.usuario === D.cfg.visador || p.usuario === D.cfg.firmante) { aviso('No se puede quitar: es el visador o el firmante.'); break; } D.cfg.personas.splice(i, 1); await guardaCfg('personas del SGD'); break; }
      case 'gem-test': await probarGemini(); break;
      case 'act-buscar': await buscarActualizacion(); break;
      case 'tras': await traspasar(); break;
      case 'tras-dev': await aplicarTraspaso(R.base, 'Fin de suplencia', 'manual'); break;
      case 'usr-add': { const u = valor('#nuUsr').toLowerCase(), r = valor('#nuRol'), nm = valor('#nuNom').toUpperCase(); if (!/^[a-z][a-z0-9._-]{2,}$/.test(u)) { aviso('Escribe el usuario del SGD tal cual (letras minúsculas, sin espacios; ej. cnunez).'); break; } if (R.roles[u]) { aviso('Ese usuario ya existe (' + esc(R.roles[u]) + ').'); break; } R.roles[u] = r; if (nm) R.nombres[u] = nm; R.admins = Object.keys(R.roles).filter(k => R.roles[k] === 'Administrador'); await guardarRoles('usuario ' + u + ' agregado'); break; }
      case 'usr-del': { const u = d.u; if (u === R.titular) { aviso('No se puede quitar al titular. Traspasa primero el trabajo.'); break; } if (R.roles[u] === 'Administrador' && R.admins.length === 1) { aviso('Debe quedar al menos un administrador.'); break; } delete R.roles[u]; delete R.nombres[u]; R.admins = Object.keys(R.roles).filter(k => R.roles[k] === 'Administrador'); await guardarRoles('usuario ' + u + ' quitado'); break; }
    }
  });
  root.addEventListener('change', async e => {
    const t = e.target, d = t.dataset;
    if (cambioAnalisis(t, d)) { /* listo */ }
    else if (d.selb !== undefined) { t.checked ? S.selB.add(+d.selb) : S.selB.delete(+d.selb); render(); }
    else if (d.filtro) { S.filtro = d.filtro; render(); }
    else if (t.name === 'apT') { const c = $('#apCat'); if (c) c.style.display = t.value === 'sug' ? 'block' : 'none'; const x = $('#apTxt'); if (x) x.placeholder = t.value === 'sug' ? 'Motivo de la sugerencia (opcional)…' : 'Escribe tu observación…'; }
    else if (d.flt) { S.tb.f[d.flt] = t.value; S.tb.pag = 1; renderTabla(); }
    else if (t.id === 'tbTam') { S.tb.tam = t.value; S.tb.pag = 1; renderTabla(); }
    else if (d.cf !== undefined) { setPath(D.cfg, d.cf, t.type === 'checkbox' ? t.checked : t.value); await guardaCfg(); }
    else if (d.catT !== undefined) { D.catalogo[+d.catT].t = t.value; await guardaCfg('catálogo de análisis'); }
    else if (d.catR !== undefined) { const x = D.catalogo[+d.catR], v = t.value; delete x.dejar; delete x.porDefecto; if (v === 'dejar') x.dejar = true; if (v === 'inf') { D.catalogo.forEach(y => { if (y !== x && y.porDefecto === 'INF') delete y.porDefecto; }); x.porDefecto = 'INF'; } await guardaCfg('catálogo de análisis'); }
    else if (d.catK !== undefined) { D.catalogo[+d.catK].kw = t.value.split(',').map(s => s.trim()).filter(Boolean); await guardaCfg('catálogo de análisis'); }
    else if (d.lc) { L[d.lc] = t.value; saveL(); toast('Guardado en este computador'); }
    else if (d.ua) { R[d.ua] = Math.max(1, +t.value || 1); await guardarRoles('alertas'); iniciarSondeo(); }
    else if (d.nom !== undefined) { R.nombres[d.nom] = t.value.trim().toUpperCase(); await guardarRoles('nombre de ' + d.nom); }
    else if (d.rol) { const u = d.rol, nuevo = t.value; if (u === R.admins[0] && nuevo !== 'Administrador' && R.admins.length === 1) { aviso('Debe quedar al menos un administrador.'); render(); return; } R.roles[u] = nuevo; R.admins = Object.keys(R.roles).filter(k => R.roles[k] === 'Administrador'); await guardarRoles('rol de ' + u); }
  });
  root.addEventListener('input', e => {
    const t = e.target, d = t.dataset || {};
    if (d.up) { const a = t.selectionStart; t.value = t.value.toUpperCase(); try { t.setSelectionRange(a, a); } catch (x) { /* ok */ } }
    else if (d.flt && t.tagName === 'INPUT') { S.tb.f[d.flt] = t.value; S.tb.pag = 1; renderTabla(); }
    else if (d.cbin) cbAbrir(d.cbin, t.value, true);
    else if (d.ufil) { S.uFil = t.value; render(); const f = $('[data-ufil]'); if (f) { f.focus(); f.setSelectionRange(f.value.length, f.value.length); } }
  });
  root.addEventListener('keydown', e => {
    const t = e.target; if (t.id === 'pinIn' && e.key === 'Enter') { root.querySelector('[data-a="pin"]').click(); return; }
    if (t.dataset && t.dataset.cus !== undefined && e.key === 'Enter') { e.preventDefault(); const o = ofAct(); if (!o || !esTit()) return; const i = +t.dataset.cus; o.bloques[i].custom = t.value; if (!t.value.trim()) { o.bloques[i].st = 'sin'; pintarFilaEstado(i); return; } accionAnalisis('conf', { i }); return; }
    if (t.classList && t.classList.contains('cbIn')) {
      const box = $('#cbl-' + t.dataset.cbin), its = [...box.querySelectorAll('.cbi')], i = its.findIndex(x => x.classList.contains('hl'));
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); if (!box.classList.contains('on')) { cbAbrir(t.dataset.cbin, t.value, true); return; } const n = Math.max(0, Math.min(its.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1))); its.forEach(x => x.classList.remove('hl')); if (its[n]) { its[n].classList.add('hl'); its[n].scrollIntoView({ block: 'nearest' }); } }
      else if (e.key === 'Enter') { e.preventDefault(); const x = its[i < 0 ? 0 : i]; if (x) cbPick(x.dataset.cbpick, x.dataset.v); }
      else if (e.key === 'Escape') { cbCerrar(); t.blur(); }
    }
  });
  // globito instantáneo: solo aparece si el texto está cortado con «…» (el del navegador tarda ~1 s)
  const moverTip = e => { const tp = $('#tip'), w = tp.offsetWidth, h = tp.offsetHeight; let x = e.clientX + 14, y = e.clientY + 16; if (x + w > innerWidth - 8) x = innerWidth - w - 8; if (y + h > innerHeight - 8) y = e.clientY - h - 10; tp.style.left = Math.max(4, x) + 'px'; tp.style.top = Math.max(4, y) + 'px'; };
  root.addEventListener('mouseover', e => { const t = e.target.closest && e.target.closest('[data-tip]'); if (!t || (t.scrollHeight <= t.clientHeight + 1 && t.scrollWidth <= t.clientWidth + 1)) return; const tp = $('#tip'); tp.textContent = t.dataset.tip; tp.style.display = 'block'; moverTip(e); });
  root.addEventListener('mousemove', e => { if ($('#tip').style.display === 'block') moverTip(e); });
  root.addEventListener('mouseout', e => { if (e.target.closest && e.target.closest('[data-tip]')) $('#tip').style.display = 'none'; });
  root.addEventListener('mouseover', e => { const it = e.target.closest && e.target.closest('.cbi'); if (!it) return; it.parentNode.querySelectorAll('.cbi.hl').forEach(x => { if (x !== it) x.classList.remove('hl'); }); it.classList.add('hl'); });   // el mouse mueve la marca: nunca hay dos filas marcadas
  root.addEventListener('focusin', e => { const t = e.target; if (t.classList && t.classList.contains('cbIn')) { t.select(); cbCerrar(); cbAbrir(t.dataset.cbin, '', false); } });
  root.addEventListener('dblclick', e => { const tr = e.target.closest && e.target.closest('[data-fila]'); if (tr) { S.sel = tr.dataset.fila; S.exp = S.sel; S.lec = { exp: S.sel }; S.vista = 'lectura'; render(); } });
  root.addEventListener('dragover', e => { const z = e.target.closest && e.target.closest('#drop'); if (z) { e.preventDefault(); z.classList.add('sobre'); } });
  root.addEventListener('dragleave', e => { const z = e.target.closest && e.target.closest('#drop'); if (z) z.classList.remove('sobre'); });
  root.addEventListener('drop', e => { const z = e.target.closest && e.target.closest('#drop'); if (z) { e.preventDefault(); z.classList.remove('sobre'); cargarArchivos(e.dataTransfer.files); } });
  $('#fileIn').addEventListener('change', async e => { const f = [...e.target.files]; e.target.value = ''; if (f.length) await cargarArchivos(f); });
  $('#bMin').onclick = minimizar;
  $('#bCerrar').onclick = async () => { if (S.dirty) { const r = await dialogo('PNF', 'Hay cambios sin guardar. ¿Quieres guardarlos antes de cerrar?', [['g', 'Guardar y cerrar', true], ['n', 'No guardar'], ['c', 'Cancelar']]); if (r === 'c') return; if (r === 'g') await guardar('al cerrar'); else S.dirty = false; } minimizar(); };
  window.addEventListener('beforeunload', e => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && $('#win').classList.contains('on') && esTit()) { e.preventDefault(); guardar('manual', true); } });
  (function () {
    const f = $('#fab'); let d = null;
    if (L.fab) { f.style.left = L.fab.x + 'px'; f.style.top = L.fab.y + 'px'; f.style.right = 'auto'; f.style.bottom = 'auto'; }
    f.addEventListener('pointerdown', e => { const r = f.getBoundingClientRect(); d = { dx: e.clientX - r.left, dy: e.clientY - r.top, x0: e.clientX, y0: e.clientY, mov: false }; try { f.setPointerCapture(e.pointerId); } catch (x) { /* ok */ } });
    f.addEventListener('pointermove', e => { if (!d) return; if (Math.abs(e.clientX - d.x0) + Math.abs(e.clientY - d.y0) > 4) d.mov = true; if (!d.mov) return; f.style.left = Math.min(Math.max(e.clientX - d.dx, 0), innerWidth - f.offsetWidth) + 'px'; f.style.top = Math.min(Math.max(e.clientY - d.dy, 0), innerHeight - f.offsetHeight) + 'px'; f.style.right = 'auto'; f.style.bottom = 'auto'; });
    f.addEventListener('pointerup', () => { if (!d) return; const mov = d.mov; d = null; if (mov) { L.fab = { x: parseInt(f.style.left, 10), y: parseInt(f.style.top, 10) }; saveL(); } else restaurar(); });
  })();

  /* ═══════════════ 9 · Lectura de oficios (etapa 2) ═══════════════ */
  const PDFLIB = () => (typeof pdfjsLib !== 'undefined' ? pdfjsLib : (W.pdfjsLib || null));
  /* Lector normal de tablas del oficio (pdf.js): usa las líneas del dibujo para saber qué texto va en cada celda. */
  function crearLector(pdfjsLib) {
    const OPS = pdfjsLib.OPS;
    const PAINT = new Set([OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke, OPS.stroke, OPS.closeStroke, OPS.closeFillStroke, OPS.closeEOFillStroke]);
    function segmentos(ol, H) {
      const hor = [], ver = []; let ctm = [1, 0, 0, 1, 0, 0]; const st = [];
      const mul = (M, o) => [M[0] * o[0] + M[1] * o[2], M[0] * o[1] + M[1] * o[3], M[2] * o[0] + M[3] * o[2], M[2] * o[1] + M[3] * o[3], M[4] * o[0] + M[5] * o[2] + o[4], M[4] * o[1] + M[5] * o[3] + o[5]];
      const ap = (x, y) => [ctm[0] * x + ctm[2] * y + ctm[4], ctm[1] * x + ctm[3] * y + ctm[5]];
      for (let i = 0; i < ol.fnArray.length; i++) {
        const f = ol.fnArray[i], a = ol.argsArray[i];
        if (f === OPS.save) st.push(ctm.slice());
        else if (f === OPS.restore) ctm = st.pop() || [1, 0, 0, 1, 0, 0];
        else if (f === OPS.transform) ctm = mul(a, ctm);
        else if (f === OPS.constructPath && PAINT.has(ol.fnArray[i + 1])) {
          const ops = a[0], args = a[1]; let k = 0;
          for (const o of ops) {
            if (o === OPS.rectangle) {
              const p1 = ap(args[k], args[k + 1]), p2 = ap(args[k] + args[k + 2], args[k + 1] + args[k + 3]); k += 4;
              const x0 = Math.min(p1[0], p2[0]), x1 = Math.max(p1[0], p2[0]), y0 = H - Math.max(p1[1], p2[1]), y1 = H - Math.min(p1[1], p2[1]), w = x1 - x0, h = y1 - y0;
              if (h < 2.5 && w > 8) hor.push({ y: (y0 + y1) / 2, x0, x1 }); else if (w < 2.5 && h > 5) ver.push({ x: (x0 + x1) / 2, y0, y1 });
            } else if (o === OPS.moveTo || o === OPS.lineTo) k += 2; else if (o === OPS.curveTo) k += 6; else if (o === OPS.curveTo2 || o === OPS.curveTo3) k += 4;
          }
        }
      }
      return { hor, ver };
    }
    async function datosPagina(page) {
      const vp = page.getViewport({ scale: 1 }), H = vp.height, tc = await page.getTextContent();
      const items = tc.items.filter(i => i.str && i.str.trim()).map(i => ({ s: i.str, x: i.transform[4], y: H - i.transform[5], w: i.width, h: i.height || 8 }));
      const seg = segmentos(await page.getOperatorList(), H);
      // el N° del oficio y la fecha vienen en campos de formulario (no en el texto normal): se leen aparte
      const anots = []; try { (await page.getAnnotations()).forEach(a => { const v = a.fieldValue != null ? a.fieldValue : ((a.contentsObj && a.contentsObj.str) || a.contents || ''); String(Array.isArray(v) ? v.join('\n') : v).split(/\n+/).map(x => x.trim()).filter(Boolean).forEach(x => anots.push(x)); }); } catch (e) { /* sin campos */ }
      return { items, hor: seg.hor, ver: seg.ver, H, W: vp.width, anots };
    }
    const agrupa = (vals, tol) => { const s = vals.slice().sort((a, b) => a - b), out = []; for (const v of s) { if (out.length && v - out[out.length - 1].m < tol) { const g = out[out.length - 1]; g.n++; g.m = (g.m * (g.n - 1) + v) / g.n; } else out.push({ m: v, n: 1 }); } return out.map(g => g.m); };
    // une los trozos de texto de una celda en líneas legibles
    function textoCelda(items) {
      const L = [];
      items.slice().sort((a, b) => a.y - b.y || a.x - b.x).forEach(i => { const l = L.find(l => Math.abs(l.y - i.y) < 2.2); if (l) l.it.push(i); else L.push({ y: i.y, it: [i] }); });
      return L.sort((a, b) => a.y - b.y).map(l => { let t = '', fin = null; l.it.sort((a, b) => a.x - b.x).forEach(i => { if (fin !== null && i.x - fin > 0.9) t += ' '; t += i.s; fin = i.x + i.w; }); return t.trim(); });
    }
    const unir = lineas => lineas.join(' ').replace(/\s+([,.;:])/g, '$1').replace(/\s{2,}/g, ' ').trim();
    function bloquesDePagina(pg) {
      // el encabezado se reconoce por la palabra PLACA (la palabra FECHA a veces viene partida); su altura de referencia es la de FECHA/LUGAR/HORA
      const heads0 = pg.items.filter(i => i.s.trim() === 'PLACA').map(pl => { const ref = pg.items.find(j => /^(FECHA|LUGAR|HORA)$/.test(j.s.trim()) && Math.abs(j.y - pl.y) < 10); return { y: ref ? ref.y : pl.y + 4, s: 'FECHA' }; }).sort((a, b) => a.y - b.y);
      const heads = heads0.filter((h, i) => !i || h.y - heads0[i - 1].y > 12);   // sin el 3.er parámetro de filter: Prototype.js 1.6 del SGD no lo entrega
      const out = [];
      const tops = heads.map(h => { const ar = pg.hor.filter(s => s.y < h.y - 3 && s.y > h.y - 30).map(s => s.y); return ar.length ? Math.max(...ar) : h.y - 14; });
      heads.forEach((h, hi) => {
        // columnas: las líneas verticales que cruzan la franja del encabezado
        const vs = pg.ver.filter(v => v.y0 <= h.y - 2 && v.y1 >= h.y + 2 && v.x > 60 && v.x < pg.W - 30);
        const xs = agrupa(vs.map(v => v.x), 2.5);
        if (xs.length < 7) return;
        const cols = []; for (let i = 0; i < xs.length - 1 && cols.length < 6; i++) cols.push({ x0: xs[i], x1: xs[i + 1] });
        const arriba = pg.hor.filter(s => s.y < h.y - 3 && s.x0 <= cols[0].x0 + 3 && s.x1 >= cols[0].x1 - 3).map(s => s.y);
        const yTop = tops[hi], sig = heads[hi + 1], yFin = sig ? tops[hi + 1] + 0.5 : 1e9;
        const bordes = cols.map(c => agrupa(pg.hor.filter(s => s.y >= yTop - 1 && s.y <= yFin && Math.min(s.x1, c.x1) - Math.max(s.x0, c.x0) > 0.6 * (c.x1 - c.x0)).map(s => s.y), 1.5));
        const celdas = bordes.map(b => { const o = []; for (let i = 0; i < b.length - 1; i++) o.push({ y0: b[i], y1: b[i + 1] }); return o; });
        const dentro = (c, cel) => pg.items.filter(i => { const cx = i.x + i.w / 2, cy = i.y - i.h * 0.3; return cx >= c.x0 && cx <= c.x1 && cy >= cel.y0 - 0.5 && cy <= cel.y1 + 0.5; });
        const tit = unir(textoCelda(dentro(cols[5], celdas[5][0] || { y0: yTop, y1: yTop + 20 })));
        const filas = [];
        (celdas[4] || []).slice(1).forEach(rc => {
          const cy = (rc.y0 + rc.y1) / 2, get = ci => { const cel = (celdas[ci] || []).find(x => cy >= x.y0 - 0.3 && cy <= x.y1 + 0.3); return cel ? { cel, it: dentro(cols[ci], cel) } : { cel: null, it: [] }; };
          const f = get(0), l = get(1), ho = get(2), fo = get(3), pl = get(4), tx = get(5);
          const lin = it => textoCelda(it);
          const fila = { fecha: lin(f.it).join('').replace(/\s/g, ''), lugar: unir(lin(l.it)), hora: lin(ho.it).join(' ').replace(/:\s+/g, ':').replace(/\s+/g, ' ').trim(), folio: unir(lin(fo.it)), placa: lin(pl.it).join('').replace(/[\s-]/g, '').toUpperCase(), texto: unir(lin(tx.it)), _y: cy };
          if (fila.placa === 'SINPPU') fila.placa = 'SIN PPU';
          if (fila.placa || fila.hora || fila.fecha) filas.push(fila);
        });
        out.push({ titulo: tit, filas, y: h.y });
      });
      return out;
    }
    async function leerTablas(doc) {
      const res = [];
      for (let p = 1; p <= doc.numPages; p++) { const pg = await datosPagina(await doc.getPage(p)); bloquesDePagina(pg).forEach(b => { b.pagina = p; res.push(b); }); }
      return res;
    }
    return { leerTablas, datosPagina, bloquesDePagina, textoCelda, unir };
  }
  let _lec = null;
  function lector() { if (_lec) return _lec; const lib = PDFLIB(); if (!lib) throw new Error('No se cargó el lector de PDF (pdf.js). Revisa la conexión a internet y que Tampermonkey tenga permiso para descargar los componentes del script.'); _lec = crearLector(lib); return _lec; }
  const lim = x => String(x == null ? '' : x).replace(/\s+/g, ' ').trim();
  const normFecha = x => { const m = lim(x).replace(/\//g, '-').match(/(\d{1,2})-(\d{1,2})-(\d{4})/); return m ? pad(+m[1]) + '-' + pad(+m[2]) + '-' + m[3] : lim(x); };
  const normHora = x => lim(x).replace(/:\s+/g, ':').replace(/\s+[aA]\s+/g, ' A ');
  const normPlaca = x => { const p = lim(x).toUpperCase().replace(/[\s-]/g, ''); return p === 'SINPPU' ? 'SIN PPU' : p; };
  const EXP_RE = /^E\d{1,9}\/\d{4}$/;
  const FECHA_RE = /^\d{2}\/\d{2}\/\d{4}$/;
  const TIPOS_DOC = ['Carta', 'Memorándum', 'Oficio', 'Ordinario', 'Correo'];
  const ART = { 'Carta': 'la', 'Memorándum': 'el', 'Oficio': 'el', 'Ordinario': 'el', 'Correo': 'el' };
  const tdoc = o => o.tipoDoc || 'Carta';
  const dmy = t => { const p = String(t || '').split('/').map(Number); return new Date(p[2], p[1] - 1, p[0]); };
  const diasDe = o => { const f = o.llegada || o.fechaDoc; return f ? diasDesde(f) : 0; };
  const nbl = o => (o.bloques || []).length;
  const nok = o => (o.bloques || []).filter(b => b.st === 'ok').length;
  const plazoTxt = n => n === 0 ? 'hoy' : n === 1 ? '1 día' : n + ' días';

  /* — textos del encabezado y datos fijos del oficio (reglas, no IA) — */
  function lineasPagina(pg) {
    const Ls = []; pg.items.slice().sort((a, b) => a.y - b.y || a.x - b.x).forEach(i => { const l = Ls.find(l => Math.abs(l.y - i.y) < 2.2); if (l) l.it.push(i); else Ls.push({ y: i.y, it: [i] }); });
    return Ls.sort((a, b) => a.y - b.y).map(l => { let t = '', fin = null; l.it.sort((a, b) => a.x - b.x).forEach(i => { if (fin !== null && i.x - fin > 0.9) t += ' '; t += i.s; fin = i.x + i.w; }); return t.trim(); }).filter(Boolean).concat(pg.anots || []);
  }
  function leerExpediente(pags) {
    for (let i = pags.length - 1; i >= 0; i--) {
      const Ls = lineasPagina(pags[i]);
      for (let k = 0; k < Ls.length; k++) {
        const m = Ls[k].match(/(?:^|\s)(E\d{1,9}\/\d{4})(?:\s|$)/); if (!m) continue;
        const mismo = Ls[k].match(/(\d{6,9})\s+E\d{1,9}\/\d{4}/);   // el N° interno del SGD va justo antes del expediente (en la misma línea o en la anterior)
        const prev = mismo ? mismo[1] : (Ls[k - 1] && /^\d{6,9}$/.test(Ls[k - 1]) ? Ls[k - 1] : '');
        return { exp: m[1], idSgd: prev };
      }
    }
    return null;
  }
  // El número propio del oficio es el de la línea «Carta N° … Santiago, fecha» (página 1) o el del pie de la última página; el ANT. también cita otros documentos.
  function leerDocumento(pags) {
    const textos = pags.map(pg => lineasPagina(pg).join(' '));
    const re = /(Carta|Memor[aá]ndum|Oficio|Ordinario|Correo)\s*N\s*[°º.]*\s*([0-9][0-9.]*\/\d{4}(?:\s+[A-Z]{2,}(?![a-záéíóúñ]))?)\s*(?:Santiago\s*,\s*(\d{2}\/\d{2}\/\d{4}))?/g;
    let mejor = null;
    textos.forEach((t, pi) => { if (pi !== 0 && pi !== textos.length - 1) return; for (const m of t.matchAll(re)) { const sc = (m[3] ? 3 : 0) + (pi === textos.length - 1 ? 2 : 0); if (!mejor || sc > mejor.sc) mejor = { sc, m }; } });
    if (!mejor) return { tipoDoc: '', numero: '', fechaDoc: '' };
    const m = mejor.m, f = m[3] || ((textos[0].match(/Santiago\s*,\s*(\d{2}\/\d{2}\/\d{4})/) || [])[1]) || '';
    return { tipoDoc: /^memor/i.test(m[1]) ? 'Memorándum' : m[1][0].toUpperCase() + m[1].slice(1).toLowerCase(), numero: lim(m[2]), fechaDoc: f };
  }
  const nombreCargo = a => { const t = a.map(x => x.trim()).filter(Boolean); return t.length ? t[0] + (t.length > 1 ? ' — ' + t.slice(1).join(' ') : '') : ''; };
  function leerEncabezado(lineas) {
    const F = { ant: [], mat: [], de: [], para: [], cc: [] }; let cur = null;
    for (const ln of lineas) {
      if (/^_{8,}/.test(ln)) break;
      const m = ln.match(/^(ANT|MAT|DE|PARA|CC)\s*\.?\s*:\s*(.*)$/i);
      if (m) { cur = m[1].toLowerCase(); if (m[2]) F[cur].push(m[2]); continue; }
      if (cur) F[cur].push(ln);
    }
    const j = a => lim(a.join(' '));
    return { ant: j(F.ant), mat: j(F.mat), de: nombreCargo(F.de), para: nombreCargo(F.para), cc: nombreCargo(F.cc) };
  }
  function filasNormal(pags) {
    const out = []; let t = 0;
    pags.forEach((pg, pi) => { lector().bloquesDePagina(pg).forEach(b => { t++; b.filas.forEach(f => out.push({ t, fecha: normFecha(f.fecha), lugar: lim(f.lugar), titulo: lim(b.titulo), hora: normHora(f.hora), folio: lim(f.folio), placa: normPlaca(f.placa), texto: lim(f.texto) })); }); });
    return out;
  }

  /* — lectura con Gemini — */
  const PROMPT_IA = `Eres un extractor de datos de oficios de la Subsecretaría de Transportes de Chile. Lee el PDF adjunto y devuelve SOLO el JSON pedido, copiando el texto tal como está escrito (sin resumir ni corregir).
Datos del oficio: tipoDoc (Carta, Memorándum, Oficio u Ordinario tal como dice el documento), numero (por ejemplo "6058/2026 PNF"), fechaDoc (dd/mm/aaaa, junto a «Santiago,»), ant (texto completo del ANT.), mat (texto del MAT.), de, para y cc (nombre y cargo, en una sola línea cada uno).
Tablas: cada tabla del oficio que empieza con el encabezado FECHA | LUGAR | HORA | FOLIO N° | PLACA PATENTE | <última columna> es un elemento de "tablas". "titulo" es el texto EXACTO del encabezado de la última columna (por ejemplo INFRACCIÓN, DETALLE DE LA FISCALIZACIÓN u OBSERVACIONES). Cada fila de la tabla es una placa: devuélvela con fecha (dd-mm-aaaa), lugar, hora, folio, placa y texto (el contenido de la última columna). Si una celda combinada abarca varias filas (fecha, lugar, hora, folio o texto), repite su contenido en cada fila. La hora se copia tal cual (incluidos rangos como «10:30 A 10:45»). Si el folio está vacío usa ""; si dice otra cosa (por ejemplo «Vehículo Particular») cópialo. La placa va en mayúsculas sin guiones («SIN PPU» tal cual). No omitas ninguna fila ni ninguna tabla y no inventes datos. No incluyas el número de expediente (E…/AAAA).`;
  const IA_SCHEMA = { type: 'OBJECT', properties: { tipoDoc: { type: 'STRING' }, numero: { type: 'STRING' }, fechaDoc: { type: 'STRING' }, ant: { type: 'STRING' }, mat: { type: 'STRING' }, de: { type: 'STRING' }, para: { type: 'STRING' }, cc: { type: 'STRING' },
    tablas: { type: 'ARRAY', items: { type: 'OBJECT', properties: { titulo: { type: 'STRING' }, filas: { type: 'ARRAY', items: { type: 'OBJECT', properties: { fecha: { type: 'STRING' }, lugar: { type: 'STRING' }, hora: { type: 'STRING' }, folio: { type: 'STRING' }, placa: { type: 'STRING' }, texto: { type: 'STRING' } }, required: ['fecha', 'lugar', 'hora', 'folio', 'placa', 'texto'] } } }, required: ['titulo', 'filas'] } } }, required: ['tablas'] };
  function bytesAB64(u8) { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); }
  async function leerConIA(bytes) {
    const body = { contents: [{ role: 'user', parts: [{ inline_data: { mime_type: 'application/pdf', data: bytesAB64(bytes) } }, { text: PROMPT_IA }] }], generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: IA_SCHEMA } };
    const r = await http({ method: 'POST', url: 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(L.modelo) + ':generateContent?key=' + encodeURIComponent(L.clave), headers: { 'Content-Type': 'application/json' }, data: jsonStr(body) });
    if (r.status !== 200) { let m = ''; try { m = JSON.parse(r.text).error.message; } catch (e) { /* sin detalle */ } throw new Error('Gemini respondió ' + r.status + (m ? ': ' + m.slice(0, 140) : '')); }
    let o; try { const j = JSON.parse(r.text); o = JSON.parse(j.candidates[0].content.parts.map(p => p.text || '').join('')); } catch (e) { throw new Error('Gemini devolvió una respuesta que no se pudo interpretar'); }
    const filas = []; (o.tablas || []).forEach((t, ti) => (t.filas || []).forEach(f => filas.push({ t: ti + 1, fecha: normFecha(f.fecha), lugar: lim(f.lugar), titulo: lim(t.titulo), hora: normHora(f.hora), folio: lim(f.folio), placa: normPlaca(f.placa), texto: lim(f.texto) })));
    if (!filas.length) throw new Error('Gemini no encontró filas en las tablas');
    return { filas, tipoDoc: lim(o.tipoDoc), numero: lim(o.numero), fechaDoc: lim(o.fechaDoc), ant: lim(o.ant), mat: lim(o.mat), de: lim(o.de), para: lim(o.para), cc: lim(o.cc) };
  }

  /* — comparar las dos lecturas y armar los bloques de análisis — */
  const claveFila = f => f.fecha + '|' + f.placa;
  function comparar(a, b) {
    if (!a || !b) return { ok: !!(a || b), soloA: [], soloB: [], menores: [] };
    const cnt = (arr, fn) => { const m = {}; arr.forEach(f => { const k = fn(f); (m[k] = m[k] || []).push(f); }); return m; };
    const A = cnt(a.filas, claveFila), B = cnt(b.filas, claveFila), soloA = [], soloB = [], menores = [];
    new Set([...Object.keys(A), ...Object.keys(B)]).forEach(k => {
      const x = A[k] || [], y = B[k] || [];
      for (let i = 0; i < Math.max(x.length, y.length); i++) {
        if (!y[i]) soloA.push(x[i]); else if (!x[i]) soloB.push(y[i]);
        else if (x[i].hora !== y[i].hora || (x[i].folio !== y[i].folio && x[i].folio && y[i].folio)) menores.push(k.replace('|', ' · ') + ': hora/folio distintos (IA «' + x[i].hora + ' ' + x[i].folio + '» · normal «' + y[i].hora + ' ' + y[i].folio + '»)');
      }
    });
    return { ok: !soloA.length && !soloB.length, soloA, soloB, menores };
  }
  function agrupar(filas) {
    const out = []; let ult = null;
    filas.forEach(f => { const k = [f.t, f.fecha, f.lugar, f.titulo, f.texto].join('\u0001'); const it = { h: f.hora, f: f.folio, p: f.placa }; if (ult && ult.k === k) ult.b.items.push(it); else { ult = { k, b: { fecha: f.fecha, lugar: f.lugar, tipo: f.titulo, texto: f.texto, items: [it], cat: null, custom: '', st: 'sin' } }; out.push(ult.b); } });
    return out;
  }
  const filasDe = e => (e.usa === 'ia' && e.ia ? e.ia.filas : e.normal.filas);
  const filasDeOficio = o => (o.bloques || []).flatMap(b => b.items.map(i => ({ fecha: b.fecha, lugar: b.lugar, titulo: b.tipo, texto: b.texto, hora: i.h, folio: i.f, placa: i.p })));

  /* — leer un PDF — */
  const leerArchivo = file => limpio(() => leerArchivo0(file));
  async function leerArchivo0(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const tarea = PDFLIB().getDocument({ data: bytes.slice(0), disableFontFace: true, isEvalSupported: false, verbosity: 0 }), doc = await tarea.promise;
    const pags = []; try { for (let p = 1; p <= doc.numPages; p++) pags.push(await lector().datosPagina(await doc.getPage(p))); } finally { try { await tarea.destroy(); } catch (x) { /* ya cerrado */ } }
    const ex = leerExpediente(pags), dc = leerDocumento(pags), enc = leerEncabezado(lineasPagina(pags[0]));
    const normal = { filas: filasNormal(pags) };
    let ia = null, errIA = '';
    if (L.motor === 'gemini') { if (!L.clave) errIA = 'no hay clave de Gemini (Configuración › Lectura y expediente)'; else { try { ia = await leerConIA(bytes); } catch (e) { errIA = e.message || String(e); } } } else errIA = 'motor normal elegido en Configuración';
    const pick = (k) => (ia && ia[k]) || enc[k] || '';
    const e = { id: Date.now() + '-' + Math.random().toString(36).slice(2, 6), archivo: file.name, tam: bytes.length, pags: doc.numPages, bytes, exp: ex ? ex.exp : '', idSgd: ex ? ex.idSgd : '',
      tipoDoc: dc.tipoDoc || '',   // si la regla no lo reconoce queda vacío: se pide elegirlo antes de agregar (no se adivina)
      numero: dc.numero || (ia && ia.numero) || '', fechaDoc: dc.fechaDoc || (ia && ia.fechaDoc) || '',
      ant: pick('ant'), mat: pick('mat'), de: pick('de'), para: pick('para'), cc: pick('cc'), ia, normal, usa: ia ? 'ia' : 'normal', errIA, modelo: ia ? L.modelo : '' };
    e.cmp = comparar(ia, ia ? normal : null);
    return e;
  }
  /* — PDFs guardados en este computador (no van al respaldo) — */
  const PDFS = {
    _db: null,
    abrir() { return new Promise((res, rej) => { if (this._db) return res(this._db); const q = indexedDB.open('pnf-pdfs', 1); q.onupgradeneeded = () => q.result.createObjectStore('p'); q.onsuccess = () => { this._db = q.result; res(q.result); }; q.onerror = () => rej(q.error); }); },
    async guardar(k, bytes) { const db = await this.abrir(); return new Promise((res, rej) => { const q = db.transaction('p', 'readwrite').objectStore('p').put(bytes.buffer ? bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) : bytes, k); q.onsuccess = () => res(); q.onerror = () => rej(q.error); }); },
    async leer(k) { const db = await this.abrir(); return new Promise((res, rej) => { const q = db.transaction('p').objectStore('p').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); }
  };

  /* — cola de oficios leídos que esperan revisión — */
  S.cola = []; S.leyendo = null; S.lec = null; S.sel = '';
  S.tb = { f: { exp: '', of: '', llego: '', asunto: '', para: '', ana: 'todos', memo: 'todos', est: 'todos' }, ord: { c: 'auto', d: 1 }, pag: 1, tam: 10 };
  async function cargarArchivos(files) {
    const lista = [...files].filter(f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
    if (!lista.length) return aviso('Elige uno o más archivos PDF de oficios.');
    if (!esTit()) return agregarOficioSupervisor(lista);
    try {
      for (let i = 0; i < lista.length; i++) {
        S.leyendo = { i: i + 1, n: lista.length, nombre: lista[i].name }; toast('Leyendo ' + lista[i].name + ' (' + (i + 1) + ' de ' + lista.length + ')…'); render();
        ocupado('Leyendo el oficio…', lista[i].name + (lista.length > 1 ? ' (' + (i + 1) + ' de ' + lista.length + ')' : ''));
        try { const e = await leerArchivo(lista[i]); S.cola.push(e); if (!entradaPend()) S.lec = { pend: e.id }; }
        catch (x) { libre(); S.leyendo = null; render(); await aviso('No se pudo leer <b>' + esc(lista[i].name) + '</b>:<br>' + esc(errTxt(x))); }
      }
    } finally { libre(); }
    S.leyendo = null; if (S.cola.length) { S.lec = { pend: entradaPend() ? S.lec.pend : S.cola[0].id }; S.vista = 'lectura'; toast('Lectura lista: revisa el oficio y pulsa «Agregar a la lista»'); } render();
  }
  const entradaPend = () => (S.lec && S.lec.pend ? S.cola.find(c => c.id === S.lec.pend) : null);

  /* — agregar, corregir y quitar oficios — */
  function construirOficio(e, ya, por) {
    const filas = filasDe(e), bloques = agrupar(filas), nIA = e.ia ? e.ia.filas.length : null;
    const o = { exp: e.exp, idSgd: e.idSgd, tipoDoc: e.tipoDoc, numero: e.numero, fechaDoc: e.fechaDoc, llegada: ya ? ya.llegada : '', mat: e.mat, ant: e.ant, de: e.de, para: e.para, cc: e.cc, archivo: e.archivo, pags: e.pags,
      sgd: ya ? ya.sgd : 'Pendiente', sgdTxt: ya ? ya.sgdTxt : '', memoOk: false, bloques, cargadoPor: por || L.usuario, cargado: new Date().toISOString(),
      lectura: { fuente: e.usa, placasIA: nIA, placasNormal: e.normal.filas.length, ok: e.cmp.ok, difs: e.cmp.soloA.length + e.cmp.soloB.length, menores: e.cmp.menores.length, modelo: e.modelo, errIA: e.errIA, fecha: new Date().toISOString() } };
    aplicarReglas(o, false); return o;
  }
  async function incorporarOficio(e, por) {
    D.oficios[e.exp] = construirOficio(e, D.oficios[e.exp], por);
    try { await PDFS.guardar(e.exp, e.bytes); } catch (x) { toast('El oficio se agregó, pero no se pudo guardar el PDF original en este computador.'); }
  }
  async function elegirTipoDoc(e) {
    const r = await dialogo('¿Qué tipo de documento es?', `<div>PNF no pudo reconocer si <b>${esc(e.archivo)}</b> es una Carta, un Memorándum, un Oficio, un Ordinario o un Correo, y <b>no lo adivina</b>. Elígelo para poder agregarlo:</div><select id="tdSel" style="margin-top:8px;width:100%"><option value="">— elegir —</option>${TIPOS_DOC.map(t => `<option>${t}</option>`).join('')}</select>`, [['ok', 'Agregar', true], ['no', 'Cancelar']], id => id === 'ok' && !valor('#tdSel') ? 'Elige el tipo de documento.' : null);
    return r === 'ok' ? valor('#tdSel') : '';
  }
  async function agregarPend(id) {
    const e = S.cola.find(x => x.id === id); if (!e) return;
    if (!EXP_RE.test(e.exp)) { await corregirDatos(); if (!EXP_RE.test(e.exp)) return; }
    if (!TIPOS_DOC.includes(e.tipoDoc)) { const t = await elegirTipoDoc(e); if (!t) return; e.tipoDoc = t; }
    const filas = filasDe(e); if (!filas.length) { const c = await dialogo('Sin tablas', '<div>No se encontró ninguna fila en las tablas de este oficio. ¿Agregarlo igual?</div>', [['si', 'Agregar igual'], ['no', 'Cancelar', true]]); if (c !== 'si') return; }
    const ya = D.oficios[e.exp];
    if (ya) { const c = await dialogo('El oficio ya está cargado', `<div><b>${esc(e.exp)}</b> ya está en la lista (${esc(tdoc(ya))} ${esc(ya.numero)}, estado <b>${esc(ya.sgd)}</b>, análisis ${nok(ya)}/${nbl(ya)}).<br>Si lo reemplazas, se pierde el análisis que ya hiciste en él.</div>`, [['rep', 'Reemplazar la lectura'], ['no', 'Cancelar', true]]); if (c !== 'rep') return; }
    await incorporarOficio(e, e.agregadoPor);
    if (e.aporteId) { D.aporteEst = D.aporteEst || {}; D.aporteEst[e.aporteId] = 'incorporado'; try { await FS.borrar(e.archivo, 'buzon'); } catch (x) { /* ya no estaba */ } }
    S.cola = S.cola.filter(x => x.id !== id); S.sel = e.exp; S.exp = e.exp;
    const sig = S.cola[0]; S.lec = sig ? { pend: sig.id } : { exp: e.exp }; if (!sig) S.vista = 'oficios';
    await guardar('oficio ' + e.exp + ' agregado'); render();
  }
  async function descartarPend(id) {
    const e = S.cola.find(x => x.id === id); if (!e) return;
    const c = await dialogo('Descartar lectura', `<div>¿Descartar la lectura de <b>${esc(e.archivo)}</b>? El archivo no se agrega a la lista.</div>`, [['si', 'Descartar'], ['no', 'Cancelar', true]]); if (c !== 'si') return;
    if (e.aporteId) { D.aporteEst = D.aporteEst || {}; D.aporteEst[e.aporteId] = 'descartada'; try { await FS.borrar(e.archivo, 'buzon'); } catch (x) { /* ya no estaba */ } await guardar('oficio de ' + e.agregadoPor + ' descartado'); }
    S.cola = S.cola.filter(x => x.id !== id); S.lec = S.cola[0] ? { pend: S.cola[0].id } : null; if (!S.cola.length) S.vista = 'oficios'; render();
  }
  async function corregirDatos() {
    const pend = entradaPend(), o = pend || (S.lec && S.lec.exp ? D.oficios[S.lec.exp] : null); if (!o) return;
    const inp = (id, v, ph) => `<input type="text" id="${id}" value="${esc(v || '')}" placeholder="${ph || ''}" style="width:100%">`, ta = (id, v) => `<textarea id="${id}" style="min-height:44px">${esc(v || '')}</textarea>`;
    const html = (pend ? `<label class="lb">Expediente (E + número / año)</label>${inp('cdExp', pend.exp, 'E155797/2026')}` : `<div>Expediente: <b>${esc(o.exp)}</b></div>`) +
      `<div class="row" style="margin-top:6px"><div><label class="lb">Tipo</label><select id="cdTipo">${!o.tipoDoc && pend ? '<option value="" selected>— elegir —</option>' : ''}${TIPOS_DOC.map(t => `<option ${(o.tipoDoc || (pend ? '' : 'Carta')) === t ? 'selected' : ''}>${t}</option>`).join('')}</select></div><div><label class="lb">N°</label>${inp('cdNum', o.numero, '6058/2026 PNF')}</div><div><label class="lb">Fecha (dd/mm/aaaa)</label>${inp('cdFec', o.fechaDoc, '01/10/2026')}</div></div>` +
      `<label class="lb">MAT.</label>${ta('cdMat', o.mat)}<label class="lb">ANT.</label>${ta('cdAnt', o.ant)}<label class="lb">DE</label>${ta('cdDe', o.de)}<label class="lb">PARA</label>${ta('cdPara', o.para)}<label class="lb">CC</label>${ta('cdCc', o.cc)}`;
    const r = await dialogo('Corregir datos del oficio', html, [['ok', 'Guardar', true], ['no', 'Cancelar']], id => { if (id !== 'ok') return null; if (pend && !EXP_RE.test(valor('#cdExp'))) return 'El expediente debe ser E + número + / + año de 4 dígitos (ej. E155797/2026).'; if (pend && !valor('#cdTipo')) return 'Elige el tipo de documento (Carta, Memorándum, Oficio, Ordinario o Correo).'; if (valor('#cdFec') && !FECHA_RE.test(valor('#cdFec'))) return 'Escribe la fecha como dd/mm/aaaa.'; return null; });
    if (r !== 'ok') return;
    if (pend) o.exp = valor('#cdExp');
    o.tipoDoc = valor('#cdTipo'); o.numero = valor('#cdNum'); o.fechaDoc = valor('#cdFec'); o.mat = valor('#cdMat'); o.ant = valor('#cdAnt'); o.de = valor('#cdDe'); o.para = valor('#cdPara'); o.cc = valor('#cdCc');
    if (!pend) await guardar('datos del oficio corregidos'); render();
  }
  async function quitarOficio(ex) {
    const o = D.oficios[ex]; if (!o) return;
    const c = await dialogo('Quitar de la lista', `<div>¿Quitar <b>${esc(ex)}</b> de la lista? Se pierde su análisis. Esta acción no se puede deshacer (queda la copia de seguridad anterior en la carpeta).</div>`, [['si', 'Quitar'], ['no', 'Cancelar', true]]); if (c !== 'si') return;
    delete D.oficios[ex]; if (S.sel === ex) S.sel = ''; if (S.lec && S.lec.exp === ex) S.lec = null; await guardar('oficio ' + ex + ' quitado', true); render();
  }
  async function verPdf(ex) {
    const e = entradaPend(); let bytes = e && !ex ? e.bytes : null;
    if (!bytes) { try { bytes = await PDFS.leer(ex || (S.lec && S.lec.exp)); } catch (x) { bytes = null; } }
    if (!bytes) return aviso('El PDF original no está guardado en este computador (se guarda solo donde se cargó el oficio).');
    window.open(URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' })), '_blank');
  }

  /* — selector buscable de oficios (sirve con 500 expedientes) — */
  function cbHtml(id, val) { return `<span class="cb" data-cb="${id}"><input type="text" class="cbIn" data-cbin="${id}" value="${esc(val)}" placeholder="${id === 'selXl' ? 'Todos los expedientes (escribe para filtrar)' : 'Escribe el expediente, el N° de oficio o parte del asunto…'}" autocomplete="off" style="width:310px"><button class="vb peq" data-cbtog="${id}" title="Mostrar la lista">▾</button><div class="cbList" id="cbl-${id}"></div></span>`; }
  function cbAbrir(id, texto, buscar) {
    const q = buscar ? texto.trim().toLowerCase() : '', ofs = datos().oficios || {}, ks = Object.keys(ofs), MAX = 8, rk = e => ofs[e].sgd === 'Respondido' ? 1 : 0;
    const Lx = (q ? ks.filter(e => (e + ' ' + ofs[e].numero + ' ' + ofs[e].mat + ' ' + ofs[e].para + ' ' + tdoc(ofs[e])).toLowerCase().includes(q)) : ks.filter(e => ofs[e].sgd !== 'Respondido')).sort((a, b) => rk(a) - rk(b) || diasDe(ofs[b]) - diasDe(ofs[a]));
    let h = (id === 'selXl' && !q ? '<div class="cbi" data-cbpick="selXl" data-v="">★ Todos los expedientes</div>' : '') + Lx.slice(0, MAX).map(e => { const o = ofs[e]; return `<div class="cbi" data-cbpick="${id}" data-v="${esc(e)}"><b class="mono">${esc(e)}</b> · ${esc(tdoc(o))} ${esc(o.numero.split('/')[0])} <span class="estado ${claseEstado(o.sgd)}">${esc(o.sgd)}</span><div class="cbs">${esc(o.mat)} — ${esc(o.para)}</div></div>`; }).join('');
    h += Lx.length ? `<div class="cbf">${Lx.length > MAX ? 'Mostrando ' + MAX + ' de ' + Lx.length + (q ? ' coincidencias' : ' pendientes') + '. Sigue escribiendo para acotar.' : Lx.length + (q ? ' coincidencia(s)' : ' pendiente(s)')}${q ? '' : ' · escribe para buscar también entre los respondidos (' + ks.length + ' en total)'} · ↑↓ y Enter</div>` : '<div class="cbf">Sin coincidencias.</div>';
    const box = $('#cbl-' + id); box.innerHTML = h; box.classList.add('on');
    // una sola fila marcada: al abrir la lista, la del oficio que está abierto (si aparece); si estás escribiendo para buscar, la primera coincidencia
    const items = [...box.querySelectorAll('.cbi')], cur = buscar ? null : (($('.cbIn[data-cbin="' + id + '"]') || {}).value || ''), marca = (cur !== null && items.find(x => x.dataset.v === cur)) || items[0]; if (marca) marca.classList.add('hl');
  }
  const cbCerrar = () => $$('.cbList.on').forEach(x => x.classList.remove('on'));
  function cbPick(id, v) { cbCerrar(); if (id === 'selXl') { S.xlExp = v; render(); return; } if (S.dirty) guardar('al cambiar de oficio'); S.sel = v; S.lec = { exp: v }; S.exp = v; S.selB.clear(); render(); }
  /* — tabla de oficios: filtros, orden y páginas — */
  function filasOf() {
    const ofs = datos().oficios || {}, f = S.tb.f, q = (a, b) => !b || String(a).toLowerCase().includes(b.toLowerCase());
    let Lx = Object.keys(ofs).map(e => { const o = ofs[e]; return { e, o, n: diasDe(o), oficio: tdoc(o) + ' ' + (o.numero || '').split('/')[0], ok: nok(o), tot: nbl(o) }; });
    Lx = Lx.filter(r => q(r.e, f.exp) && q(r.oficio, f.of) && q(r.o.llegada || r.o.fechaDoc, f.llego) && q(r.o.mat, f.asunto) && q(r.o.para, f.para) && (f.ana === 'todos' || (f.ana === 'completos' ? r.ok === r.tot : r.ok < r.tot)) && (f.memo === 'todos' || (f.memo === 'Sí') === !!r.o.memoOk) && (f.est === 'todos' || (f.est === 'activos' ? r.o.sgd !== 'Respondido' : r.o.sgd === f.est)));
    const rank = r => r.o.sgd === 'Devuelto' ? 0 : r.o.sgd === 'Por despachar' ? 1 : r.o.sgd === 'Pendiente' ? 2 : 3, col = S.tb.ord.c, d = S.tb.ord.d;
    const key = { exp: r => r.e, of: r => r.oficio, llego: r => +dmy(r.o.llegada || r.o.fechaDoc), lleva: r => r.n, asunto: r => r.o.mat, para: r => r.o.para, ana: r => (r.tot ? r.ok / r.tot : 0), memo: r => r.o.memoOk ? 1 : 0, est: r => rank(r) }[col];
    Lx.sort(key ? (a, b) => (key(a) > key(b) ? 1 : key(a) < key(b) ? -1 : 0) * d : (a, b) => rank(a) - rank(b) || (rank(a) === 3 ? +dmy(b.o.fechaDoc) - +dmy(a.o.fechaDoc) : b.n - a.n));
    return Lx;
  }
  const corta = f => { const p = String(f || '').split('/'); return p.length === 3 ? p[0] + '/' + p[1] + '/' + p[2].slice(2) : (f || '—'); };
  function renderTabla() {
    const tb = $('#tbBody'); if (!tb) return;
    const Lx = filasOf(), tam = S.tb.tam === 'Todos' ? (Lx.length || 1) : +S.tb.tam, pags = Math.max(1, Math.ceil(Lx.length / tam)); S.tb.pag = Math.min(S.tb.pag, pags);
    const ini = (S.tb.pag - 1) * tam, vis = Lx.slice(ini, ini + tam);
    tb.innerHTML = vis.length ? vis.map(r => { const o = r.o, cs = claseEstado(o.sgd), pc = r.tot ? Math.round(r.ok * 100 / r.tot) : 0, med = o.cargadoPor && o.cargadoPor !== R.titular ? ` title="Agregado por ${esc(o.cargadoPor)}"` : '';
      return `<tr class="${r.e === S.sel ? 'sel ' : ''}${o.sgd === 'Devuelto' ? 'dv' : ''}" data-fila="${esc(r.e)}"><td class="exp"${med}>${med ? '<span class="tag">＋</span> ' : ''}${esc(r.e)}</td><td>${esc(r.oficio)}</td><td>${corta(o.llegada || o.fechaDoc)}</td><td class="c">${o.sgd === 'Respondido' ? '—' : `<span class="estado ${r.n > 30 ? 'e-r' : r.n > 14 ? 'e-a' : ''}">${r.n} d</span>`}</td><td data-tip="${esc(o.mat)}">${esc(o.mat)}</td><td data-tip="${esc(o.para)}">${esc(o.para)}</td><td class="c"><span class="pbm"><i style="width:${pc}%"></i></span> ${r.ok}/${r.tot}</td><td class="c">${o.memoOk ? 'Sí' : '—'}</td><td class="c"><span class="estado ${cs}">${esc(o.sgd)}</span></td></tr>`; }).join('')
      : '<tr><td colspan="9" style="text-align:center;padding:16px;color:#666">' + (Object.keys(datos().oficios || {}).length ? 'No hay oficios que coincidan con los filtros.' : 'Todavía no hay oficios cargados.') + '</td></tr>';
    $$('.sortk').forEach(e => { e.textContent = ''; }); const sk = $('#sk-' + S.tb.ord.c); if (sk) sk.textContent = S.tb.ord.d > 0 ? '▲' : '▼';
    const nb = (t, p, dis, on) => `<span class="${dis ? 'dis' : on ? 'on' : ''}"${dis ? '' : ' data-pg="' + p + '"'}>${t}</span>`, pg = S.tb.pag;
    let h = nb('«', 1, pg === 1) + nb('‹', pg - 1, pg === 1); const a = Math.max(1, Math.min(pg - 2, pags - 4)), z = Math.min(pags, a + 4);
    for (let p = a; p <= z; p++) h += nb(p, p, false, p === pg); h += nb('›', pg + 1, pg === pags) + nb('»', pags, pg === pags);
    $('#tbPag').innerHTML = `<div class="pag"><span>Mostrar</span><select id="tbTam">${[10, 25, 50, 'Todos'].map(x => `<option ${String(S.tb.tam) === String(x) ? 'selected' : ''}>${x}</option>`).join('')}</select><span>por página</span><span class="sp"></span><span>${Lx.length ? (ini + 1) + '–' + (ini + vis.length) : '0'} de ${Lx.length}</span><div class="n">${h}</div></div>`;
  }

  /* — pantalla Oficios — */
  function vOficios() {
    const t = esTit(), dt = datos(), n = Object.keys(dt.oficios || {}).length, F = S.tb.f;
    const inp = (k, ph) => `<input type="text" data-flt="${k}" placeholder="${ph}" value="${esc(F[k])}">`, sel = (k, ops) => `<select data-flt="${k}">${ops.map(o => `<option value="${o[0]}" ${F[k] === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>`;
    const tit = [['exp', 'Expediente'], ['of', 'Oficio'], ['llego', 'Llegó'], ['lleva', 'Lleva'], ['asunto', 'Asunto'], ['para', 'Para'], ['ana', 'Análisis'], ['memo', 'Memo'], ['est', 'Estado']].map(c => `<th><span class="ord" data-ord="${c[0]}">${c[1]}</span><span class="sortk" id="sk-${c[0]}"></span></th>`).join('');
    const COLS = '<col style="width:130px"><col style="width:118px"><col style="width:74px"><col style="width:60px"><col><col><col style="width:108px"><col style="width:58px"><col style="width:116px">';
    const carga = t ? frame('Cargar oficio', (S.leyendo ? `<div class="note inf">⏳ Leyendo <b>${esc(S.leyendo.nombre)}</b> (${S.leyendo.i} de ${S.leyendo.n})… puede tardar unos segundos.</div>` : '') +
      `<div class="drop" data-a="examinar" id="drop"><b>Arrastra aquí el PDF del oficio</b><br>o pulsa para elegirlo (puedes elegir varios)</div><div style="font-size:11.5px;margin-top:6px">${L.motor === 'gemini' && L.clave ? 'Se lee con Gemini y se comprueba con el lector normal del PDF. El expediente lo lee siempre una regla fija, nunca la IA.' : 'Se lee con el lector normal del PDF' + (L.motor === 'gemini' ? ' (falta la clave de Gemini en Configuración › Lectura y expediente).' : '.')}</div>` +
      (S.cola.length ? `<div style="margin-top:8px"><b>Leídos, por revisar (${S.cola.length})</b><div class="lv" style="height:auto;max-height:110px;margin-top:4px"><table>${S.cola.map(c => `<tr><td class="exp">${esc(c.exp || '(sin expediente)')}</td><td>${esc(c.archivo)}</td><td>${c.cmp.ok && TIPOS_DOC.includes(c.tipoDoc) ? '<span class="estado e-v">cuadra</span>' : '<span class="estado e-a">revisar</span>'}</td><td>${btn('Revisar', 'data-a="lec-ir" data-id="' + esc(c.id) + '"', 'peq def')}</td></tr>`).join('')}</table></div></div>` : '')) :
      frame('Agregar oficio', aporta() ? (S.leyendo ? `<div class="note inf">⏳ Revisando <b>${esc(S.leyendo.nombre)}</b>…</div>` : '') + '<div class="drop" data-a="examinar" id="drop"><b>Arrastra aquí el PDF del oficio</b><br>o pulsa para elegirlo</div><div style="font-size:11.5px;margin-top:6px">Antes de agregar, PNF actualiza los datos y revisa que no esté repetido. El PDF queda en el buzón y ' + esc(R.titular) + ' lo lee y lo incorpora solo.</div>' : '<div class="vacio">Tu usuario no tiene rol para agregar oficios.</div>');
    const estado = frame('Estado de PNF', '<dl class="kv">' + dl('Tú en PNF', L.usuario ? '<b>' + esc(L.usuario) + '</b> · ' + etiquetaRol() : '<span class="estado e-r">sin identificar</span> ' + btn('Identificarme', 'data-a="ident"', 'peq')) + dl('Titular actual', '<b>' + esc(R.titular) + '</b> desde ' + esc(R.desde) + (R.hasta ? ' hasta ' + esc(R.hasta) : '')) +
      dl('Datos', t ? 'revisión <b>' + D.rev + '</b> · guardado ' + cuando(D.guardado) : 'revisión <b>' + (V.rev || 0) + '</b> de ' + esc(R.titular) + ' · ' + cuando(V.guardado)) + dl('Carpeta', { ok: '<span class="estado e-v">conectada</span> ' + esc(S.carpetaNombre), sin: '<span class="estado e-a">sin conectar</span>', permiso: '<span class="estado e-a">falta confirmar el permiso</span> ' + esc(S.carpetaNombre) }[S.carpeta]) + dl('Bandeja del SGD', datos().chkAt ? cuando(datos().chkAt) + (datos().chkRes ? ' · ' + datos().chkRes.enBandeja + ' en la bandeja' : '') : '<span class="estado e-a">sin verificar</span>') + dl('Último respaldo', cuando(L.ultResp)) + dl('Versión de PNF', VERSION) + '</dl>' + (t ? '<div class="row" style="margin-top:6px">' + btn('Verificar bandeja', 'data-a="verif"', 'peq') + '</div>' : ''));
    return `<div class="g2">${carga}${estado}</div>` + frame('Oficios cargados (' + n + ')', `<div class="tw"><table class="t" id="tbOf"><colgroup>${COLS}</colgroup><thead><tr>${tit}</tr><tr class="fl"><th>${inp('exp', 'buscar…')}</th><th>${inp('of', 'N°')}</th><th>${inp('llego', 'dd/mm')}</th><th></th><th>${inp('asunto', 'buscar…')}</th><th>${inp('para', 'buscar…')}</th><th>${sel('ana', [['todos', 'todos'], ['sin', 'sin terminar'], ['completos', 'completos']])}</th><th>${sel('memo', [['todos', 'todos'], ['Sí', 'Sí'], ['No', 'No']])}</th><th>${sel('est', [['todos', 'Todos'], ['activos', 'Pend. + Devueltos'], ['Por despachar', 'Por despachar'], ['Pendiente', 'Pendiente'], ['Devuelto', 'Devuelto'], ['Respondido', 'Respondido']])}</th></tr></thead><tbody id="tbBody"></tbody></table></div><div id="tbPag"></div>` +
      `<div class="row" style="margin-top:8px">${btn('Abrir seleccionado', 'data-a="abrir-sel"')}${t ? btn('Quitar de la lista', 'data-a="quitar-sel"') : ''}<span class="sp"></span><span style="font-size:11.5px">Clic en un título = ordenar · escribe bajo el título para filtrar · doble clic en una fila = abrir. Orden inicial: devueltos, pendientes (los más antiguos primero) y respondidos.</span></div>`);
  }

  /* — pantalla Lectura del oficio — */
  function vLectura() {
    const pend = entradaPend(), ex = pend ? '' : (S.lec && S.lec.exp) || '', o = pend || (ex ? D.oficios[ex] : null) || (datos().oficios || {})[S.sel];
    if (!o) return barraOficio() + frame('Lectura del oficio', '<div class="vacio">Elige un oficio arriba, o carga uno nuevo desde «Oficios».</div>');
    const filas = pend ? filasDe(pend) : filasDeOficio(o), t = esTit();
    const L2 = x => `<div data-tip="${esc(x)}" style="max-height:34px;overflow:hidden">${esc(x)}</div>`;
    const expTxt = pend ? pend.exp : o.exp;
    let ctrl;
    if (pend) {
      const c = pend.cmp, ia = pend.ia, nn = pend.normal.filas.length;
      ctrl = (!pend.tipoDoc ? '<div class="note mal">⚠ No se reconoció el tipo de documento (Carta, Memorándum, Oficio, Ordinario o Correo). PNF no lo adivina: al agregar se te pedirá elegirlo.</div>' : '') + (ia ? (c.ok ? `<div class="note ok">✔ La lectura cuadra: las filas que devolvió la IA coinciden con las que el lector normal encuentra en el PDF.</div>` : `<div class="note mal">⚠ No cuadra: ${c.soloA.length} fila(s) las ve solo la IA y ${c.soloB.length} solo el lector normal. Revisa antes de agregar.</div>`) : `<div class="note warn">ℹ Se usó solo el lector normal${pend.errIA ? ' (' + esc(pend.errIA) + ')' : ''}. Revisa las tablas con cuidado.</div>`) +
        `<dl class="kv" style="grid-template-columns:150px minmax(0,1fr)">${dl('Lectura en uso', pend.usa === 'ia' ? '<b>IA</b> · ' + esc(pend.modelo) : '<b>Lector normal</b>')}${dl('Filas (lector normal)', nn)}${dl('Filas (IA)', ia ? ia.filas.length : '—')}${dl('Bloques de análisis', agrupar(filas).length)}${dl('Expediente', 'lo lee una regla fija, <b>nunca la IA</b>')}${c.menores.length ? dl('Diferencias menores', c.menores.length + ' (hora o folio)') : ''}</dl>` +
        (!c.ok && ia ? `<div class="lv" style="height:auto;max-height:110px;margin-top:6px"><table>${c.soloA.slice(0, 20).map(f => `<tr><td class="estado e-a">solo IA</td><td>${esc(f.fecha)}</td><td class="mono">${esc(f.placa)}</td><td>${esc(f.hora)}</td></tr>`).join('')}${c.soloB.slice(0, 20).map(f => `<tr><td class="estado e-r">solo normal</td><td>${esc(f.fecha)}</td><td class="mono">${esc(f.placa)}</td><td>${esc(f.hora)}</td></tr>`).join('')}</table></div>` : '') +
        `<div class="row" style="margin-top:8px">${ia ? btn('Usar lectura IA', 'data-a="lec-usa" data-m="ia"', pend.usa === 'ia' ? 'peq def' : 'peq') + btn('Usar lectura normal', 'data-a="lec-usa" data-m="normal"', pend.usa === 'normal' ? 'peq def' : 'peq') : ''}${L.clave && L.motor === 'gemini' ? btn(ia ? 'Releer con IA' : 'Intentar con IA', 'data-a="lec-ia"', 'peq') : ''}</div>`;
    } else {
      const l = o.lectura || {};
      ctrl = (l.fuente === 'ia' ? (l.ok ? '<div class="note ok">✔ Cuando se agregó, la IA y el lector normal coincidían.</div>' : `<div class="note warn">⚠ Cuando se agregó había ${l.difs} fila(s) distinta(s) entre la IA y el lector normal; se revisaron al agregar.</div>`) : '<div class="note warn">ℹ Se leyó solo con el lector normal.</div>') +
        `<dl class="kv" style="grid-template-columns:150px minmax(0,1fr)">${dl('Lectura usada', l.fuente === 'ia' ? '<b>IA</b> · ' + esc(l.modelo || '') : '<b>Lector normal</b>')}${dl('Filas (lector normal)', l.placasNormal != null ? l.placasNormal : '—')}${dl('Filas (IA)', l.placasIA != null ? l.placasIA : '—')}${dl('Bloques de análisis', nbl(o))}${dl('Cargado', esc(cuando(o.cargado)) + ' por ' + esc(o.cargadoPor || '—'))}</dl>`;
    }
    const sets = pend && pend.cmp && pend.ia ? { a: new Set(pend.cmp.soloA.map(claveFila)), b: new Set(pend.cmp.soloB.map(claveFila)) } : { a: new Set(), b: new Set() };
    const tablaLec = `<div class="lv" style="max-height:330px"><table><tr><th>Fecha</th><th>Lugar</th><th>Hora</th><th>Folio</th><th>Placa</th><th>Tipo</th><th>Texto del oficio</th></tr>${filas.map(f => { const mal = sets.a.has(claveFila(f)) || sets.b.has(claveFila(f)); return `<tr${mal ? ' style="background:#ffe0e0"' : ''}><td class="mono" style="white-space:nowrap">${esc(f.fecha)}</td><td>${esc(f.lugar)}</td><td class="mono" style="white-space:nowrap">${esc(f.hora)}</td><td class="mono">${esc(f.folio)}</td><td class="mono"><b>${esc(f.placa)}</b></td><td>${esc(f.titulo)}</td><td>${esc(f.texto)}</td></tr>`; }).join('')}</table></div>`;
    const bar = pend ? `<div class="note inf" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><span>👁 Revisando <b>${esc(pend.archivo)}</b> (${S.cola.indexOf(pend) + 1} de ${S.cola.length}). Todavía no está en la lista.</span><span class="sp"></span>${btn('Agregar a la lista', 'data-a="lec-agregar" data-id="' + esc(pend.id) + '"', 'def')}${btn('Descartar', 'data-a="lec-descartar" data-id="' + esc(pend.id) + '"')}</div>` : '';
    return barraOficio() + bar + '<div class="g2">' +
      frame('Identificación', `<dl class="kv"><dt>Expediente</dt><dd><span class="mono" style="font-size:15px"><b>${esc(expTxt || '—')}</b></span> ${EXP_RE.test(expTxt) ? '<span class="estado e-v">formato válido</span>' : '<span class="estado e-r">no se encontró</span>'}</dd><dt>Documento</dt><dd>${pend && !o.tipoDoc ? '<span class="estado e-r">tipo sin reconocer</span>' : esc(tdoc(o))} N° ${esc(o.numero)} · ${esc(o.fechaDoc)}</dd><dt>MAT.</dt><dd>${L2(o.mat)}</dd><dt>ANT.</dt><dd>${L2(o.ant)}</dd><dt>DE</dt><dd>${L2(o.de)}</dd><dt>PARA</dt><dd>${L2(o.para)}</dd><dt>CC</dt><dd>${L2(o.cc)}</dd><dt>Archivo</dt><dd>${esc(o.archivo)} · ${o.pags} págs.</dd></dl><div class="row" style="margin-top:8px">${t ? btn('Corregir datos…', 'data-a="lec-corregir"', 'peq') : ''}${btn('Ver PDF original', 'data-a="ver-pdf"', 'peq')}</div>`) +
      frame('Control de lectura', ctrl) + '</div>' + frame('Tablas leídas (' + filas.length + ' filas)', tablaLec) +
      `<div class="pie" style="display:flex;gap:8px;margin-top:10px"><button class="vb" data-go="oficios">← Oficios</button><span class="sp"></span>${pend ? '' : btn('Siguiente: Análisis técnico →', 'data-a="lec-analisis"', 'def')}</div>`;
  }

  /* ═══════════════ 10 · Análisis técnico (etapa 3) ═══════════════ */
  const catalogo = () => datos().catalogo || [];
  const exAct = () => S.exp || S.sel || '';
  const ofAct = () => (datos().oficios || {})[exAct()];
  function sugerir(b) {
    const h = (b.texto + ' ' + b.items.map(i => i.p).join(' ')).toLowerCase(), cat = catalogo();
    for (const c of cat) if ((c.kw || []).some(k => k && h.includes(k.toLowerCase()))) return c.id;
    const d = cat.find(c => c.porDefecto === (/^INFRACC/i.test(b.tipo) ? 'INF' : b.tipo));
    return d ? d.id : null;
  }
  function aplicarReglas(o, soloNoConfirmadas) {
    (o.bloques || []).forEach(b => {
      if (soloNoConfirmadas && b.st === 'ok') return;
      const id = sugerir(b), c = catalogo().find(x => x.id === id);
      if (!c || c.dejar) { b.cat = null; b.st = 'sin'; b.custom = ''; } else { b.cat = id; b.st = 'sug'; b.custom = ''; }
    });
  }
  const textoAn = b => b.cat === 'custom' ? b.custom : ((catalogo().find(c => c.id === b.cat) || {}).t || '');
  const conteo = o => { const rs = o.bloques || []; return { n: rs.length, ok: rs.filter(r => r.st === 'ok').length, sug: rs.filter(r => r.st === 'sug').length, sin: rs.filter(r => r.st === 'sin').length }; };
  const nPlacas = o => (o.bloques || []).reduce((a, b) => a + b.items.filter(i => i.p).length, 0);
  S.filtro = 'todas'; S.selB = new Set(); S.foco = null;

  function barraOficio(soloGuardado) {
    const pend = soloGuardado ? null : entradaPend(), ex = pend ? pend.exp : (soloGuardado ? exAct() : ((S.lec && S.lec.exp) || exAct())), o = ex && (datos().oficios || {})[ex];
    return `<div class="row" style="margin-bottom:8px"><span>Oficio:</span>${cbHtml('selExp', ex)}${pend ? '<span class="estado e-a">pendiente de agregar</span>' : o ? '<span class="estado">' + esc(tdoc(o)) + ' ' + esc(o.numero) + '</span>' : ''}<span class="sp"></span></div>`;
  }
  const roBanner = () => (!esTit() && L.usuario ? `<div class="ro">👁 Modo supervisión: estás viendo los datos de <b>${esc(R.titular)}</b> en solo lectura. ${aporta() ? 'Puedes dejar observaciones o sugerencias con el botón 💬 de cada bloque.' : 'Tu usuario no tiene rol para dejar aportes.'}</div>` : '');
  const clsAn = b => b.st === 'ok' ? 'okc' : b.st === 'sug' ? 'sug' : 'sinc';
  const estCel = (b, i, dis) => (b.st === 'ok' ? 'Confirmado' : b.st === 'sug' ? btn('Confirmar', 'data-a="conf" data-i="' + i + '"' + (dis || ''), 'peq') + '<div style="font-size:11px">' + (b.cat === 'custom' ? 'texto propio: falta confirmar (o pulsa Enter)' : 'sugerido por regla') + '</div>' : '<b>Sin asignar</b>') + apExtra(i);
  // actualiza solo las celdas de estado de una fila, sin redibujar la pantalla (así no se pierde el clic ni el foco)
  function pintarFilaEstado(i) { const o = ofAct(), b = o && o.bloques[i], tr = $('#vista tr[data-row="' + i + '"]'); if (!b || !tr) return; const td = tr.querySelectorAll('td'), a = td[td.length - 2], e = td[td.length - 1]; a.className = clsAn(b); e.className = clsAn(b); e.innerHTML = estCel(b, i, ''); }
  function vAnalisis() {
    const o = ofAct();
    if (!o) return roBanner() + barraOficio(true) + frame('Análisis técnico', '<div class="vacio">Elige un oficio arriba o en «Oficios».</div>');
    const t = esTit(), c = conteo(o), pct = c.n ? Math.round(c.ok * 100 / c.n) : 0, dis = t ? '' : ' disabled';
    const vis = o.bloques.map((b, i) => [b, i]).filter(([b]) => S.filtro === 'todas' || b.st === S.filtro);
    const optCat = sel => '<option value="">— elegir —</option>' + catalogo().filter(x => !x.dejar).map(x => `<option value="${esc(x.id)}" ${x.id === sel ? 'selected' : ''}>${esc(x.t)}</option>`).join('') + `<option value="custom" ${sel === 'custom' ? 'selected' : ''}>✎ Texto propio…</option>`;
    const filas = vis.map(([b, i]) => {
      const cls = clsAn(b), pl = b.items.filter(x => x.p).map(x => `<span class="mono">${esc(x.p)}</span>`).join(' · ');
      const hint = b.st === 'sin' && !b.cat ? '<div style="color:#900;font-size:11px;margin-top:2px">Sin análisis sugerido: elige uno de la lista (o escribe un texto propio).</div>' : '';
      return `<tr data-row="${i}"><td><input type="checkbox" data-selb="${i}" ${S.selB.has(i) ? 'checked' : ''} ${dis}></td><td style="white-space:nowrap">${esc(b.fecha)}</td><td class="txt">${esc(b.lugar)}</td><td>${esc(b.tipo)}</td><td class="txt ancho">${esc(b.texto)}</td><td style="min-width:110px">${pl}<div style="font-size:11px">${b.items.filter(x => x.p).length} placa(s)</div></td>` +
        `<td class="${cls}" style="min-width:300px"><div class="an"><span>${b.cat === 'custom' ? '✎ Texto propio…' : textoAn(b) ? esc(textoAn(b)) : '<span style="color:#777">— elegir —</span>'}</span><select data-catb="${i}" ${dis}>${optCat(b.cat)}</select></div>${b.cat === 'custom' ? `<input type="text" data-cus="${i}" value="${esc(b.custom)}" placeholder="Escribe el análisis…" style="width:100%;margin-top:3px" ${dis}>` : ''}${hint}</td>` +
        `<td class="${cls}">${estCel(b, i, dis)}</td></tr>`;
    }).join('');
    const op = (k, tx, n) => `<label style="margin-right:12px"><input type="radio" name="flt" data-filtro="${k}" ${S.filtro === k ? 'checked' : ''}> ${tx} (${n})</label>`;
    return roBanner() + barraOficio(true) + frame('Avance', `<div class="row"><span><b>${c.ok} de ${c.n}</b> bloques confirmados</span><div class="pb"><i style="width:${pct}%"></i></div><span class="sp"></span>${btn('Volver a sugerir', 'data-a="re"' + dis)}${btn('Confirmar sugeridas (' + c.sug + ')', 'data-a="conf-sug"' + (c.sug && t ? '' : ' disabled'))}</div><div style="margin-top:6px">Mostrar: ${op('todas', 'Todos', c.n)}${op('sin', 'Sin asignar', c.sin)}${op('sug', 'Sugeridos', c.sug)}${op('ok', 'Confirmados', c.ok)}</div>`) +
      frame('Bloques del oficio', `<div class="row" style="margin-bottom:6px"><span><b>${S.selB.size}</b> seleccionado(s)</span><select id="bulkCat" ${dis}>${optCat('').replace('<option value="custom">✎ Texto propio…</option>', '')}</select>${btn('Asignar a seleccionados', 'data-a="bulk"' + (S.selB.size && t ? '' : ' disabled'), 'peq')}<span class="sp"></span>${t ? btn('Editar el catálogo de análisis…', 'data-a="ir-cat"', 'peq') : ''}</div><div class="lv" style="height:520px"><table><tr><th></th><th>Fecha</th><th>Lugar</th><th>Tipo</th><th>Texto del oficio</th><th>Placas</th><th>Análisis técnico</th><th>Estado</th></tr>${filas}</table></div>`) +
      `<div class="pie" style="display:flex;gap:8px;align-items:center;margin-top:10px"><button class="vb" data-go="lectura">← Lectura del oficio</button><span class="sp"></span><span style="color:#900">${c.ok === c.n ? '' : 'Faltan ' + (c.n - c.ok) + ' bloque(s) por confirmar'}</span><button class="vb def" data-go="memo" ${c.ok === c.n ? '' : 'disabled'}>Continuar → Memo</button></div>`;
  }
  // un solo cambio no guarda (solo marca «sin guardar»); las acciones masivas sí guardan
  function enfocar(i, que) { S.foco = { i, que }; }
  async function accionAnalisis(a, d) {
    const o = ofAct(); if (!o || !esTit()) return;
    const i = +d.i;
    if (a === 'conf') { const b = o.bloques[i]; b.st = 'ok'; let n = o.bloques.findIndex((r, k) => k > i && r.st !== 'ok'); if (n < 0) n = o.bloques.findIndex((r, k) => k < i && r.st !== 'ok'); S.foco = n >= 0 ? { i: n, que: o.bloques[n].st === 'sug' ? 'conf' : 'sel' } : { i, que: 'sel' }; marcar(); render(); }
    else if (a === 'conf-sug') { o.bloques.forEach(b => { if (b.st === 'sug') b.st = 'ok'; }); await guardar('análisis confirmado'); }
    else if (a === 're') { aplicarReglas(o, true); await guardar('sugerencias recalculadas con el catálogo actual'); }
    else if (a === 'bulk') { const v = valor('#bulkCat'); if (!v) { await aviso('Elige primero un análisis en la lista.'); return; } S.selB.forEach(k => { const b = o.bloques[k]; if (b) { b.cat = v; b.st = 'ok'; b.custom = ''; } }); S.selB.clear(); await guardar('análisis asignado a varios bloques'); }
    else if (a === 'ir-cat') { S.vista = 'config'; S.cfgTab = 'cat'; render(); }
  }
  function cambioAnalisis(t, d) {
    const o = ofAct(); if (!o || !esTit()) return false;
    if (d.catb !== undefined) { const b = o.bloques[+d.catb]; b.cat = t.value || null; b.st = !t.value || t.value === 'custom' ? 'sin' : 'ok'; if (t.value !== 'custom') b.custom = ''; enfocar(+d.catb, t.value === 'custom' ? 'cus' : 'sel'); marcar(); render(); return true; }
    if (d.cus !== undefined) {   // al salir del campo solo se guarda el texto (queda por confirmar); Enter o «Confirmar» lo confirman y pasan a la siguiente fila
      const b = o.bloques[+d.cus]; if (b.st === 'ok' && b.custom === t.value) return true;
      b.custom = t.value; b.st = !t.value.trim() ? 'sin' : b.st === 'ok' ? 'ok' : 'sug'; marcar(); pintarFilaEstado(+d.cus); return true; }
    return false;
  }

  /* ═══════════════ 11 · Aportes de supervisores y avisos (carpeta compartida) ═══════════════ */
  S.aportes = []; const apCache = {};
  const aporta = () => !esTit() && !!L.usuario && rolDe(L.usuario) !== 'Espectador';
  const estAp = id => (datos().aporteEst || {})[id] || 'nuevo';
  const idAp = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  async function leerAportes() {
    if (S.carpeta !== 'ok' || !FS.h) return;
    const nombres = []; try { for await (const [n, x] of FS.h.entries()) if (x.kind === 'file' && /^aportes-.+\.json$/.test(n)) nombres.push(n); } catch (e) { return; }
    const out = [];
    for (const n of nombres) {
      const u = n.slice(8, -5); if (!esTit() && u !== L.usuario) continue;
      try {
        const f = await (await FS.h.getFileHandle(n)).getFile(), c = apCache[n]; let j;
        if (c && f.lastModified !== undefined && c.t === f.lastModified && c.s === f.size) j = c.j; else { j = JSON.parse(await f.text()); apCache[n] = { t: f.lastModified, s: f.size, j }; }
        (arr(j.items) || []).forEach(it => out.push(Object.assign({}, it, { de: u })));
      } catch (e) { /* archivo a medio escribir: se lee en el próximo ciclo */ }
    }
    S.aportes = out;
  }
  async function escribirAporte(item) {
    if (S.carpeta !== 'ok') throw new Error('La carpeta de respaldo no está conectada.');
    const n = 'aportes-' + L.usuario + '.json'; let j = { v: 1, usuario: L.usuario, items: [] };
    try { const t = await FS.leer(n); if (t) j = JSON.parse(t); } catch (e) { /* archivo nuevo */ }
    j.items = arr(j.items); if (!Array.isArray(j.items)) j.items = [];
    j.items.push(item); await FS.escribir(n, jsonStr(j)); delete apCache[n]; await leerAportes();
  }
  const aportesDe = (ex, i) => S.aportes.filter(a => a.exp === ex && a.fila === i && (a.tipo === 'obs' || a.tipo === 'sug') && estAp(a.id) === 'nuevo');
  function apExtra(i) {
    const ex = exAct(), Lx = aportesDe(ex, i);
    return (aporta() ? `<div style="margin-top:3px">${btn('💬', 'data-a="ap-nuevo" data-i="' + i + '" title="Dejar una observación o sugerir otro análisis"', 'peq')}</div>` : '') + (Lx.length ? `<div style="font-size:11px;margin-top:3px">💬 ${Lx.map(a => esc(a.de)).join(', ')}</div>` : '');
  }
  const tipoAp = a => a.tipo === 'oficio' ? '➕ Oficio agregado' : a.tipo === 'obs' ? '💬 Observación' : a.tipo === 'err' ? '🚩 Error del memo' : '✎ Sugerencia';
  function frameAportes() {
    const Lx = S.aportes.filter(a => (a.tipo === 'obs' || a.tipo === 'sug') && estAp(a.id) === 'nuevo'); if (!Lx.length || !esTit()) return '';
    return frame('Aportes de supervisión que esperan tu decisión (' + Lx.length + ')', `<div class="lv" style="height:auto;max-height:170px"><table><tr><th>De</th><th>Oficio · bloque</th><th>Aporte</th><th>Cuándo</th><th></th></tr>${Lx.map(a => `<tr><td>${esc(a.de)}</td><td class="mono">${esc(a.exp)} · ${a.fila + 1}</td><td>${a.tipo === 'obs' ? '💬 ' + esc(a.txt) : '✎ Sugiere: «' + esc(textoCat(a.cat)) + '»'}</td><td>${esc(cuando(a.hora))}</td><td style="white-space:nowrap">${btn('Revisar…', 'data-a="apr" data-id="' + esc(a.id) + '"', 'peq def')}${a.tipo === 'obs' ? btn('Visto', 'data-a="apv" data-id="' + esc(a.id) + '"', 'peq') : btn('Aplicar', 'data-a="apa" data-id="' + esc(a.id) + '"', 'peq') + btn('Descartar', 'data-a="apd" data-id="' + esc(a.id) + '"', 'peq')}</td></tr>`).join('')}</table></div><div style="font-size:11.5px;margin-top:6px">«Revisar…» muestra el bloque completo con el contexto. «Aplicar» deja ese análisis como <b>sugerido</b>; tú lo confirmas. Nada se cambia sin tu decisión.</div>`);
  }
  const textoCat = id => ((catalogo().find(c => c.id === id) || {}).t || '');
  async function marcarAporte(id, est, motivo) { D.aporteEst = D.aporteEst || {}; D.aporteEst[id] = est; await guardar(motivo); }
  async function abrirApv(id) {
    const a = S.aportes.find(x => x.id === id); if (!a) return; const o = D.oficios[a.exp], r = o && o.bloques[a.fila], dd = (k, v) => `<dt>${k}</dt><dd>${v}</dd>`;
    const base = dd('De', `<b>${esc(a.de)}</b> · ${esc(cuando(a.hora))}`) + dd('Oficio', `<span class="mono">${esc(a.exp)}</span>${o ? ' · ' + esc(tdoc(o)) + ' N° ' + esc(o.numero) : ' (ya no está en la lista)'}`);
    let h, bt;
    if (a.tipo === 'err') { h = `<dl class="kv" style="grid-template-columns:120px minmax(0,1fr)">${base}${dd('Dónde', esc(a.donde))}${dd('Qué dice', '<b>' + esc(a.txt) + '</b>')}${dd('Urgencia', a.urg ? '<span class="estado e-r">Corregir antes de enviar</span>' : 'Puede esperar')}</dl><div class="note" style="margin-top:8px">Si vale para todos los memos, corrígelo en Configuración › Memo. Si es solo para este (por ejemplo una subrogancia), usa «Cambiar el PARA solo de este memo…» en la pestaña Memo.</div>`; bt = [['memo', 'Ir al memo'], ['corr', 'Ya lo corregí', true], ['x', 'Cerrar']]; }
    else if (!r) { h = `<dl class="kv" style="grid-template-columns:120px minmax(0,1fr)">${base}${dd('Aporte', esc(a.txt || ''))}</dl><div class="note warn">El bloque ya no existe (se volvió a leer el oficio).</div>`; bt = [['desc', 'Descartar', true], ['x', 'Cerrar']]; }
    else {
      const items = r.items.map(i => `${esc(i.h)} · folio ${esc(i.f)} · <span class="mono">${esc(i.p || '—')}</span>`).join('<br>');
      h = `<dl class="kv" style="grid-template-columns:120px minmax(0,1fr)">${base}${dd('Bloque', (a.fila + 1) + ' de ' + o.bloques.length)}${dd('Fecha', esc(r.fecha))}${dd('Lugar', esc(r.lugar))}${dd('Tipo', esc(r.tipo))}${dd('Texto del oficio', esc(r.texto))}${dd('Placas', items)}${dd('Análisis actual', (r.st === 'ok' ? '<span class="estado e-v">Confirmado</span> ' : r.st === 'sug' ? '<span class="estado e-a">Sugerido por regla</span> ' : '<span class="estado e-r">Sin asignar</span> ') + (textoAn(r) ? esc(textoAn(r)) : '—'))}${a.tipo === 'sug' ? dd('Sugiere ' + esc(a.de), '<b>' + esc(textoCat(a.cat)) + '</b>' + (a.txt ? '<br><i>Motivo: ' + esc(a.txt) + '</i>' : '')) : dd('Observación', '<b>' + esc(a.txt) + '</b>')}</dl>`;
      bt = a.tipo === 'sug' ? [['apa', 'Aplicar', true], ['desc', 'Descartar'], ['bloque', 'Ir al bloque'], ['x', 'Cerrar']] : [['visto', 'Visto', true], ['bloque', 'Ir al bloque'], ['x', 'Cerrar']];
    }
    const r2 = await dialogo(a.tipo === 'err' ? 'Error reportado en el memo' : a.tipo === 'sug' ? 'Sugerencia de análisis' : 'Observación', h, bt);
    if (r2 === 'apa') { if (r) { r.cat = a.cat; r.custom = ''; r.st = 'sug'; } await marcarAporte(id, 'aplicada', 'aporte de ' + a.de + ' aplicado'); }
    else if (r2 === 'desc') await marcarAporte(id, 'descartada', 'aporte descartado');
    else if (r2 === 'visto') await marcarAporte(id, 'visto', 'aporte visto');
    else if (r2 === 'corr') await marcarAporte(id, 'corregido', 'error del memo corregido');
    else if (r2 === 'memo') { S.exp = a.exp; S.sel = a.exp; S.vista = 'memo'; }
    else if (r2 === 'bloque') { S.exp = a.exp; S.sel = a.exp; S.filtro = 'todas'; S.foco = { i: a.fila, que: 'sel' }; S.vista = 'analisis'; }
    render();
  }
  async function nuevoAporte(i) {
    const o = ofAct(), b = o.bloques[i];
    const html = `<div style="margin-bottom:6px">Oficio <b>${esc(exAct())}</b> · bloque <b>${i + 1}</b> — ${esc(b.lugar)}</div><label><input type="radio" name="apT" value="obs" checked> Observación</label> <label><input type="radio" name="apT" value="sug"> Sugerir otro análisis</label><div style="margin-top:6px"><select id="apCat" style="width:100%;display:none">${catalogo().filter(x => !x.dejar).map(x => `<option value="${esc(x.id)}">${esc(x.t)}</option>`).join('')}</select><textarea id="apTxt" style="width:100%;min-height:70px" placeholder="Escribe tu observación…"></textarea></div><div style="font-size:11.5px;margin-top:4px">No cambia nada del oficio: queda en tu buzón y ${esc(R.titular)} decide.</div>`;
    const r = await dialogo('Dejar un aporte', html, [['ok', 'Enviar a mi buzón', true], ['no', 'Cancelar']], id => id === 'ok' && $('input[name=apT]:checked').value === 'obs' && !valor('#apTxt') ? 'Escribe la observación.' : null);
    if (r !== 'ok') return;
    const tipo = $('input[name=apT]:checked').value;
    try { await escribirAporte({ id: idAp(), tipo, exp: exAct(), fila: i, cat: tipo === 'sug' ? valor('#apCat') : undefined, txt: valor('#apTxt'), hora: new Date().toISOString() }); toast('Aporte guardado en tu buzón (aportes-' + L.usuario + '.json). ' + R.titular + ' decide.'); } catch (e) { await aviso(esc(e.message)); }
    render();
  }
  async function reportarErrorMemo() {
    const html = `<div style="margin-bottom:6px">Oficio <b>${esc(exAct())}</b>. El aviso le llega a ${esc(R.titular)} en cuanto su PNF revise la carpeta (cada ${Math.round((R.seg || 60) / 60 * 10) / 10 || 1} min como máximo). La corrección la hace ${esc(R.titular)} después.</div><label class="lb">Dónde está el error</label><select id="errDonde" style="width:100%"><option>Encabezado (DE / PARA / ASUNTO)</option><option>Párrafo o texto del memo</option><option>Tablas (fechas, lugares, placas)</option><option>Análisis técnico de un bloque</option><option>Otro</option></select><label class="lb">Qué está mal</label><textarea id="apTxt" style="width:100%;min-height:70px" placeholder="Ej.: el PARA está mal: está de subrogante Claudia Núñez (cnunez)."></textarea><label style="display:block;margin-top:6px"><input type="checkbox" id="errUrg" checked> Urgente: hay que corregirlo antes de enviar el memo</label>`;
    const r = await dialogo('Reportar un error del memo', html, [['ok', 'Enviar aviso', true], ['no', 'Cancelar']], id => id === 'ok' && !valor('#apTxt') ? 'Escribe qué está mal.' : null);
    if (r !== 'ok') return;
    try { await escribirAporte({ id: idAp(), tipo: 'err', exp: exAct(), donde: valor('#errDonde'), urg: $('#errUrg').checked, txt: valor('#apTxt'), hora: new Date().toISOString() }); toast('Aviso enviado a tu buzón. ' + R.titular + ' lo verá en cuanto actualice.'); } catch (e) { await aviso(esc(e.message)); }
    render();
  }
  function alertasAportes(A) {
    if (!esTit()) return;
    const er = S.aportes.filter(a => a.tipo === 'err' && estAp(a.id) === 'nuevo');
    er.slice(0, 3).reverse().forEach(a => A.unshift(['mal', `🚩 <b>${esc(a.de)}</b> reporta un error en el memo de <b>${esc(a.exp)}</b> (${esc(a.donde)}): «${esc(a.txt)}» ${btn('Revisar…', 'data-a="apr" data-id="' + esc(a.id) + '"', 'peq def')}${btn('Ya lo corregí', 'data-a="apc" data-id="' + esc(a.id) + '"', 'peq')}`]));
    if (er.length > 3) A.splice(3, 0, ['mal', '🚩 …y ' + (er.length - 3) + ' aviso(s) más de error del memo.']);
    S.cola.filter(c => c.aporteId).forEach(c => A.push(['warn', `📥 El oficio <b>${esc(c.exp || c.archivo)}</b> que dejó <b>${esc(c.agregadoPor)}</b> en el buzón necesita tu revisión antes de entrar a la tabla (${TIPOS_DOC.includes(c.tipoDoc) ? 'la lectura con IA y la normal no coinciden' : 'PNF no reconoció el tipo de documento y no lo adivina'}). ` + btn('Revisar…', 'data-a="lec-ir" data-id="' + esc(c.id) + '"', 'peq def')]));
    Object.keys(S.buzonErr || {}).forEach(id => { const a = S.aportes.find(x => x.id === id); if (a) A.push(['mal', `⚠ No se pudo incorporar el oficio <b>${esc(a.exp)}</b> que dejó <b>${esc(a.de)}</b>. <span title="${esc(S.buzonErr[id].det)}">${esc(S.buzonErr[id].txt)}</span> PNF vuelve a intentarlo solo cada minuto; si sigue igual, avisa al administrador de PNF. ` + btn('↻ Reintentar ahora', 'data-a="act"', 'peq def')]); });
  }
  function vSup() {
    const dt = datos(), ofs = dt.oficios || {}, ks = Object.keys(ofs), act = ks.filter(e => ofs[e].sgd !== 'Respondido').sort((a, b) => diasDe(ofs[b]) - diasDe(ofs[a]));
    const nP = ks.filter(e => ofs[e].sgd === 'Pendiente').length, nD = ks.filter(e => ofs[e].sgd === 'Devuelto').length, nZ = ks.filter(e => ofs[e].sgd === 'Por despachar').length, nR = ks.length - act.length;
    const filas = act.map(e => { const o = ofs[e], c = conteo(o), n = diasDe(o), pc = c.n ? Math.round(c.ok * 100 / c.n) : 0; return `<tr data-fila="${esc(e)}"><td class="exp">${esc(e)}</td><td>${esc(tdoc(o))} ${esc(o.numero.split('/')[0])}</td><td><span class="estado ${n >= R.dias ? 'e-r' : ''}">${plazoTxt(n)}</span></td><td><span class="pbm"><i style="width:${pc}%"></i></span> ${c.ok}/${c.n}</td><td>${o.memoOk ? 'Sí' : '—'}</td><td><span class="estado ${o.sgd === 'Devuelto' ? 'e-a' : 'e-z'}">${esc(o.sgd)}</span></td><td>${esc(o.cargadoPor || R.titular)}</td></tr>`; }).join('');
    const mis = S.aportes.filter(a => a.de === L.usuario), nombre = { corregido: 'corregido por ' + R.titular, nuevo: 'esperando decisión de ' + R.titular, visto: 'visto por ' + R.titular, aplicada: 'aplicada', descartada: 'descartada', incorporado: 'ya está en la tabla de ' + R.titular };
    return roBanner() + frame('Resumen', `<div class="row"><span class="estado e-p">${nZ} por despachar</span><span class="estado e-z">${nP} pendientes</span><span class="estado e-a">${nD} devuelto(s)</span><span class="estado e-v">${nR} respondido(s)</span><span class="sp"></span><span style="font-size:11.5px">Alerta de plazo: ${R.dias} días</span></div>`) +
      frame('Avance de los oficios sin responder (' + act.length + ')', `<div class="lv" style="height:auto;max-height:250px"><table><tr><th>Expediente</th><th>Oficio</th><th>Lleva</th><th>Análisis</th><th>Memo</th><th>Estado</th><th>Cargado por</th></tr>${filas}</table></div><div style="font-size:11.5px;margin-top:6px">Doble clic en una fila para abrir el oficio (solo lectura).</div>`) +
      frame('Mis aportes (' + mis.length + ')', mis.length ? `<div class="lv" style="height:auto;max-height:170px"><table><tr><th>Qué</th><th>Oficio</th><th>Cuándo</th><th>Estado</th></tr>${mis.map(a => `<tr><td>${tipoAp(a)}</td><td class="mono">${esc(a.exp)}${a.fila !== undefined ? ' · bloque ' + (a.fila + 1) : ''}</td><td>${esc(cuando(a.hora))}</td><td>${esc(nombre[a.tipo === 'oficio' && estAp(a.id) === 'nuevo' ? 'nuevo' : estAp(a.id)] || estAp(a.id))}</td></tr>`).join('')}</table></div>` : 'Todavía no has dejado aportes. En «Análisis técnico» cada bloque tiene un botón 💬, en «Memo» puedes reportar un error y con «Agregar oficio» puedes cargar uno nuevo.');
  }

  /* ═══════════════ 12 · Memo (etapa 3): vista previa, PDF y Word con el formato oficial ═══════════════ */
  const MEMO_CSS = `
.pnfmemo .pg { background:#fff; width:816px; margin:0 auto 14px; padding:36px 0 40px; box-shadow:0 2px 8px rgba(0,0,0,.5); font-family:Verdana,Geneva,sans-serif; color:#000; box-sizing:border-box; }
.pnfmemo .pg.pdf { box-shadow:none; margin:0; height:1344px; overflow:hidden; }
.pnfmemo .pg .logo { display:block; margin:0 0 0 66px; width:116px; height:105px; }
.pnfmemo .pg table.enc { margin:22px 0 0 61px; width:648px; border-collapse:collapse; font-size:10.7px; font-weight:bold; }
.pnfmemo .pg table.enc td { border:1px dotted #555; padding:7px 9px; vertical-align:middle; } .pnfmemo .pg table.enc td.l { width:82px; }
.pnfmemo .pg .rule { margin:16px 0 0 66px; width:606px; border-top:3px solid #000; }
.pnfmemo .pg .intro { margin:18px 0 0 66px; width:606px; font:10.67px/1.4 Verdana,Geneva,sans-serif; text-align:justify; text-indent:40px; }
.pnfmemo .pg table.m { margin:14px 0 0 24px; width:768px; border-collapse:collapse; table-layout:fixed; font-size:9.33px; }
.pnfmemo .pg table.m th { background:#a2a2a2; color:#1a1a1a; font-weight:bold; font-size:9.3px; padding:4px 2px; border:1px solid #222; text-align:center; height:28px; }
.pnfmemo .pg table.m td { border:1px solid #222; padding:3px 2px; text-align:center; vertical-align:middle; line-height:1.25; overflow-wrap:anywhere; }
.pnfmemo .pg table.m td.nw { white-space:nowrap; } .pnfmemo .pg table.m td.pen { background:#fff4cf; } .pnfmemo .pg table.m td.fal { background:#ffe0e0; }
.pnfmemo .pg .cierre { margin:16px 0 0 66px; font:10.67px Verdana,Geneva,sans-serif; }`;
  const mcfg = () => datos().cfg;
  const paraDe = ex => (datos().oficios[ex] && datos().oficios[ex].paraMemo) || mcfg().para;
  const subTxt = (t, ex) => { const o = datos().oficios[ex], td = tdoc(o); return t.replace(/\{\{TIPO\}\}/g, td).replace(/\{\{NUMERO\}\}|\{\{CARTA\}\}/g, o.numero).replace(/\{\{DOC\}\}/g, (ART[td] || 'el') + ' ' + td + ' N° ' + o.numero).replace(/\ba (el|la) /g, (m, ar) => ar === 'la' ? 'a la ' : 'al '); };
  const rachas = (items, key) => { const out = new Map(); let k = 0; while (k < items.length) { let j = k; while (j + 1 < items.length && items[j + 1][key] === items[k][key]) j++; out.set(k, j - k + 1); k = j + 1; } return out; };
  const COLS_M = '<colgroup><col style="width:10.4%"><col style="width:10.4%"><col style="width:10.4%"><col style="width:10.4%"><col style="width:10.4%"><col style="width:28.2%"><col style="width:19.8%"></colgroup>';
  function tablaBloque(b, desde, hasta, marcas) {
    const items = b.items.slice(desde, hasta), n = items.length, rh = rachas(items, 'h'), rf = rachas(items, 'f');
    const cl = !marcas || b.st === 'ok' ? '' : b.st === 'sug' ? 'pen' : 'fal', an = textoAn(b) || (b.st === 'sin' ? '⚠ SIN ASIGNAR' : '');
    return '<table class="m">' + COLS_M + '<thead><tr><th>FECHA</th><th>LUGAR</th><th>HORA</th><th>FOLIO N°</th><th>PLACA<br>PATENTE</th><th>' + esc((b.tipo || '').toUpperCase()) + '</th><th>ANÁLISIS TÉCNICO</th></tr></thead><tbody>' +
      items.map((i, k) => '<tr>' + (k ? '' : `<td rowspan="${n}" class="nw">${esc(b.fecha)}</td><td rowspan="${n}">${esc(b.lugar)}</td>`) + (rh.has(k) ? `<td rowspan="${rh.get(k)}" class="nw">${esc(i.h)}</td>` : '') + (rf.has(k) ? `<td rowspan="${rf.get(k)}">${esc(i.f)}</td>` : '') + `<td>${esc(i.p)}</td>` + (k ? '' : `<td rowspan="${n}">${esc(b.texto)}</td><td rowspan="${n}" class="${cl}">${esc(an)}</td>`) + '</tr>').join('') + '</tbody></table>';
  }
  const logoMemo = () => `<img class="logo" src="${LOGO}" alt="SEREMI">`;
  function cabeceraMemo(ex) {
    const C = mcfg(), p = paraDe(ex);
    return `<table class="enc"><tr><td class="l">DE</td><td>${esc(C.de.nombre)}<br>${esc(C.de.cargo)}</td></tr><tr><td class="l">PARA</td><td>${esc(p.nombre)}<br>${esc(p.cargo)}</td></tr><tr><td class="l">ASUNTO</td><td>${esc(subTxt(C.asunto, ex))}</td></tr></table><div class="rule"></div><p class="intro">${esc(subTxt(C.intro, ex))}</p>`;
  }
  function hojaMemo(ex, marcas) {
    const o = datos().oficios[ex];
    return `<div class="pg">${logoMemo()}${cabeceraMemo(ex)}${o.bloques.map(b => tablaBloque(b, 0, b.items.length, marcas)).join('')}<p class="cierre">${esc(mcfg().cierre)}</p></div>`;
  }
  const nombreMemo = o => 'MEMO LEGAL REPORTES FISCA ' + tdoc(o) + ' ' + String(o.numero).split('/')[0];
  /* paginado: el logo se repite en cada hoja y ninguna fila queda cortada; si una tabla no cabe, sigue en la hoja siguiente con su encabezado */
  async function paginasMemo(ex) {
    const o = datos().oficios[ex], cont = document.createElement('div'); cont.className = 'pnfmemo'; cont.style.cssText = 'position:fixed;left:-20000px;top:0;z-index:-1;background:#fff';
    const st = document.createElement('style'); st.textContent = MEMO_CSS; cont.appendChild(st); document.body.appendChild(cont);
    const nueva = primera => { const pg = document.createElement('div'); pg.className = 'pg pdf'; pg.style.height = 'auto'; pg.innerHTML = logoMemo() + (primera ? cabeceraMemo(ex) : ''); cont.appendChild(pg); return pg; };
    const pages = [], MAXH = 1344; let pg = nueva(true); pages.push(pg);
    const img = pg.querySelector('img'); try { if (img && img.decode) await img.decode(); } catch (e) { /* sin logo */ } try { if (document.fonts && document.fonts.ready) await document.fonts.ready; } catch (e) { /* ok */ }
    const caben = () => pg.offsetHeight <= MAXH, nodo = h => { const d = document.createElement('div'); d.innerHTML = h; return d.firstChild; };
    let ult = null; const prueba = (b, d, h) => { ult = nodo(tablaBloque(b, d, h, false)); pg.appendChild(ult); if (caben()) return true; pg.removeChild(ult); return false; };
    for (const b of o.bloques) {
      let desde = 0; const n = b.items.length;
      while (desde < n) {
        if (prueba(b, desde, n)) { desde = n; continue; }
        let lo = 0, hi = n - desde - 1;                                   // cuántas filas caben en lo que queda de la hoja
        while (lo < hi) { const mid = Math.ceil((lo + hi) / 2); if (prueba(b, desde, desde + mid)) { pg.removeChild(ult); lo = mid; } else hi = mid - 1; }
        const vacia = pg.children.length <= (pg === pages[0] ? 4 : 1);
        if (lo >= 1) { prueba(b, desde, desde + lo); desde += lo; pg = nueva(false); pages.push(pg); }
        else if (!vacia) { pg = nueva(false); pages.push(pg); }          // aquí no cabe nada: hoja nueva y se reintenta
        else { ult = nodo(tablaBloque(b, desde, desde + 1, false)); pg.appendChild(ult); desde += 1; pg = nueva(false); pages.push(pg); }   // una sola fila que no cabe ni sola: se deja
      }
    }
    const cierre = document.createElement('p'); cierre.className = 'cierre'; cierre.textContent = mcfg().cierre; pg.appendChild(cierre); if (!caben()) { pg.removeChild(cierre); pg = nueva(false); pages.push(pg); pg.appendChild(cierre); }
    pages.forEach(p => { p.style.height = MAXH + 'px'; });
    return { cont, pages };
  }
  const memoPDF = ex => limpio(() => memoPDF0(ex));
  async function memoPDF0(ex) {
    const JS = typeof jspdf !== 'undefined' ? jspdf : W.jspdf, H2C = typeof html2canvas !== 'undefined' ? html2canvas : W.html2canvas;
    if (!JS || !H2C) throw new Error('No se cargaron los componentes para generar el PDF (jsPDF y html2canvas). Revisa la conexión y los permisos de Tampermonkey.');
    const { cont, pages } = await paginasMemo(ex);
    try {
      const pdf = new JS.jsPDF({ unit: 'pt', format: [612, 1008], orientation: 'portrait', compress: true });
      for (let k = 0; k < pages.length; k++) { const cv = await H2C(pages[k], { scale: 1.6, backgroundColor: '#ffffff', logging: false }); if (k) pdf.addPage([612, 1008], 'portrait'); pdf.addImage(cv.toDataURL('image/jpeg', 0.85), 'JPEG', 0, 0, 612, 1008, undefined, 'FAST'); }
      return { bytes: new Uint8Array(pdf.output('arraybuffer')), paginas: pages.length };
    } finally { cont.remove(); }
  }
  function memoWord(ex) {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>Memo</title><style>@page { size:8.5in 14in; margin:0.5in 0.4in; } ${MEMO_CSS.replace(/\.pnfmemo /g, '').replace(/box-shadow:[^;]+;/g, '').replace(/width:816px;/, '')} body { margin:0; }</style></head><body class="pnfmemo">${hojaMemo(ex, false)}</body></html>`;
    return new Blob(['﻿', html], { type: 'application/msword' });
  }
  function descargar(blob, nombre) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre; a.style.display = 'none'; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000); }
  S.gen = false;
  async function descargarMemo(tipo) {
    const ex = exAct(), o = datos().oficios[ex]; if (!o || S.gen) return;
    S.gen = true; toast('Generando el memo…'); render();
    try {
      if (tipo === 'pdf') { const r = await memoPDF(ex); descargar(new Blob([r.bytes], { type: 'application/pdf' }), nombreMemo(o) + '.pdf'); toast('Memo PDF generado (' + r.paginas + ' hojas)'); }
      else { descargar(memoWord(ex), nombreMemo(o) + '.doc'); toast('Memo Word generado'); }
      if (esTit()) { o.memoOk = true; await guardar('memo generado'); }
    } catch (e) { await aviso('No se pudo generar el memo: ' + esc(e.message || String(e))); }
    S.gen = false; render();
  }
  async function cambiarPara() {
    const ex = exAct(), o = datos().oficios[ex], q = paraDe(ex);
    const html = `<div style="margin-bottom:6px">Solo cambia el PARA del memo del oficio <b>${esc(ex)}</b>. Los demás memos siguen con el PARA de Configuración (<b>${esc(mcfg().para.nombre)}</b>).</div><label class="lb">PARA · nombre</label><input type="text" id="paraNom" data-up="1" value="${esc(q.nombre)}" style="width:100%;text-transform:uppercase"><label class="lb">PARA · cargo</label><input type="text" id="paraCar" data-up="1" value="${esc(q.cargo)}" style="width:100%;text-transform:uppercase"><div style="font-size:11.5px;margin-top:6px">Se escribe en MAYÚSCULAS, como en el memo oficial. Ejemplo de subrogancia: <b>CLAUDIA NÚÑEZ</b> · JEFA (S) ÁREA LEGAL - SEREMITT.</div>`;
    const r = await dialogo('Cambiar el PARA solo de este memo', html, [['ok', 'Usar en este memo', true], ['no', 'Cancelar']], id => id === 'ok' && (!valor('#paraNom') || !valor('#paraCar')) ? 'Escribe el nombre y el cargo.' : null);
    if (r !== 'ok') return;
    o.paraMemo = { nombre: valor('#paraNom').toUpperCase(), cargo: valor('#paraCar').toUpperCase() }; o.memoOk = false; await guardar('PARA solo de este memo'); render();
  }
  function vMemo() {
    const ex = exAct(), o = datos().oficios[ex];
    if (!o) return roBanner() + barraOficio(true) + frame('Memo', '<div class="vacio">Elige un oficio arriba o en «Oficios».</div>');
    const t = esTit(), c = conteo(o), lista = c.n > 0 && c.ok === c.n;
    return roBanner() + barraOficio(true) + `<div class="note ${lista ? 'ok' : 'mal'}">${lista ? '✔ Todos los bloques están confirmados: el memo está listo.' : '⚠ Faltan ' + (c.n - c.ok) + ' bloques por confirmar (amarillo = sugerido, rojo = sin asignar). Se puede ver la vista previa, pero no descargar el memo final.'}</div>` +
      `<div class="row" style="margin-bottom:8px">${btn(S.gen ? 'Generando…' : 'Descargar PDF', 'data-a="memo-pdf"' + (lista && !S.gen && t ? '' : ' disabled'), 'def')}${btn('Descargar Word', 'data-a="memo-doc"' + (lista && !S.gen && t ? '' : ' disabled'))}${btn('Editar encabezado y textos…', 'data-a="ir-memo-cfg"' + (t ? '' : ' disabled'))}${btn('Cambiar el PARA solo de este memo…', 'data-a="para-memo"' + (t ? '' : ' disabled'))}${o.paraMemo ? btn('Volver al PARA de Configuración', 'data-a="para-quitar"' + (t ? '' : ' disabled'), 'peq') + '<span class="estado e-a">PARA personalizado (solo este memo)</span>' : ''}${aporta() ? btn('🚩 Reportar un error del memo…', 'data-a="err-memo"', 'def') : ''}<span class="sp"></span><span>Nombre: ${esc(nombreMemo(o))}.pdf · sin número ni fecha (los pone el SGD)</span></div>` +
      `<div class="visor pnfmemo"><style>${MEMO_CSS}</style>${hojaMemo(ex, true)}</div><div style="font-size:11.5px;margin-top:6px">Hoja tamaño oficio (8,5 × 14 in) con el formato del memo oficial: logo SEREMI, cabecera DE / PARA / ASUNTO y tablas de columnas fijas e iguales. El logo se repite en cada hoja del PDF; los colores amarillo y rojo son solo de esta vista previa.</div>` +
      `<div class="pie" style="display:flex;gap:8px;align-items:center;margin-top:10px"><button class="vb" data-go="analisis">← Análisis técnico</button><span class="sp"></span><span style="color:#900">${lista ? '' : 'Confirma todos los bloques para poder enviar'}</span><button class="vb def" data-go="sgd" ${lista ? '' : 'disabled'}>Enviar al SGD →</button></div>`;
  }

  /* ═══════════════ 13 · Excel acumulado (etapa 4) ═══════════════ */
  S.xlExp = '';
  function filasExcel(ex) {
    const ofs = datos().oficios || {}, out = [];
    Object.keys(ofs).filter(e => !ex || e === ex).forEach(e => { const o = ofs[e]; o.bloques.forEach(b => b.items.forEach(i => { if (!i.p) return; out.push({ exp: e, numero: tdoc(o) + ' ' + o.numero, fechaDoc: o.fechaDoc, fecha: b.fecha, lugar: b.lugar, hora: i.h, folio: i.f, placa: i.p, tipo: b.tipo, texto: b.texto, analisis: textoAn(b), estado: b.st === 'ok' ? 'confirmado' : 'pendiente' }); })); });
    return out;
  }
  function vExcel() {
    const filas = filasExcel(S.xlExp), puede = rolDe(L.usuario) !== 'Espectador' || esTit();
    const h = filas.slice(0, 2000).map(f => `<tr><td>${esc(f.exp)}</td><td>${esc(f.numero)}</td><td>${esc(f.fechaDoc)}</td><td>${esc(f.fecha)}</td><td class="txt">${esc(f.lugar)}</td><td>${esc(f.hora)}</td><td>${esc(f.folio)}</td><td class="mono">${esc(f.placa)}</td><td>${esc(f.tipo)}</td><td class="txt">${esc(f.texto)}</td><td class="txt">${esc(f.analisis)}</td><td>${f.estado}</td></tr>`).join('');
    return frame('Planilla acumulada', `<div class="row"><span>Mostrar:</span>${cbHtml('selXl', S.xlExp)}<span class="estado">${filas.length} filas</span><span class="sp"></span>${btn('Exportar a Excel (.xlsx)', 'data-a="xls"' + (puede && filas.length ? '' : ' disabled'), 'def')}</div>`) +
      `<div class="lv xl" style="height:430px"><table><tr><th>Expediente</th><th>Documento N°</th><th>Fecha documento</th><th>Fecha</th><th>Lugar</th><th>Hora</th><th>Folio</th><th>Placa</th><th>Tipo</th><th>Texto del oficio</th><th>Análisis técnico</th><th>Estado análisis</th></tr>${h}</table></div>` +
      `<div class="note" style="margin-top:8px">Una fila por placa; la primera columna dice de qué expediente viene. Cada oficio que se carga <b>agrega</b> sus filas; si se vuelve a cargar el mismo expediente, reemplaza las suyas (no duplica). Todos los usuarios con rol pueden ver y descargar esta planilla.${filas.length > 2000 ? ' Se muestran las primeras 2000 filas; el archivo exportado las trae todas.' : ''}</div>`;
  }
  const exportarExcel = () => limpio(() => exportarExcel0());
  async function exportarExcel0() {
    const X = typeof XLSX !== 'undefined' ? XLSX : W.XLSX; if (!X) { await aviso('No se cargó el componente para crear el Excel (SheetJS). Revisa la conexión y los permisos de Tampermonkey.'); return; }
    const filas = filasExcel(S.xlExp); if (!filas.length) { await aviso('No hay filas para exportar.'); return; }
    const ws = X.utils.json_to_sheet(filas.map(f => ({ 'Expediente': f.exp, 'Documento N°': f.numero, 'Fecha documento': f.fechaDoc, 'Fecha': f.fecha, 'Lugar': f.lugar, 'Hora': f.hora, 'Folio': f.folio, 'Placa': f.placa, 'Tipo': f.tipo, 'Texto del oficio': f.texto, 'Análisis técnico': f.analisis, 'Estado análisis': f.estado })));
    ws['!cols'] = [14, 22, 14, 12, 40, 14, 12, 10, 24, 60, 50, 14].map(w => ({ wch: w })); const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'Acumulado');
    descargar(new Blob([X.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'PNF-acumulado' + (S.xlExp ? '-' + S.xlExp.replace('/', '-') : '') + '.xlsx'); toast('Excel exportado (' + filas.length + ' filas)');
  }

  /* ═══════════════ 14 · Oficios agregados por supervisores (buzón en la carpeta compartida) ═══════════════ */
  FS.leerBytes = async function (nombre, sub) { let d = this.h; try { if (sub) d = await d.getDirectoryHandle(sub); const f = await (await d.getFileHandle(nombre)).getFile(); return new Uint8Array(await f.arrayBuffer()); } catch (e) { if (e && (e.name === 'NotFoundError' || e.name === 'TypeMismatchError')) return null; throw e; } };
  async function agregarOficioSupervisor(lista) {
    if (!aporta()) return aviso('Tu usuario no tiene rol para agregar oficios.');
    if (S.carpeta !== 'ok') return aviso('Primero conecta la carpeta de respaldo (Configuración › Respaldo).');
    for (let i = 0; i < lista.length; i++) {
      const f = lista[i]; S.leyendo = { i: i + 1, n: lista.length, nombre: f.name }; toast('Revisando ' + f.name + '…'); render();
      ocupado('Revisando el oficio…', f.name + (lista.length > 1 ? ' (' + (i + 1) + ' de ' + lista.length + ')' : ''));
      try {
        await sincronizar();                                                    // 1) datos al día antes de revisar repetidos
        const bytes = new Uint8Array(await f.arrayBuffer()), pags = await limpio(async () => { const tarea = PDFLIB().getDocument({ data: bytes.slice(0), disableFontFace: true, isEvalSupported: false, verbosity: 0 }), doc = await tarea.promise, r = []; try { for (let p = 1; p <= doc.numPages; p++) r.push(await lector().datosPagina(await doc.getPage(p))); } finally { try { await tarea.destroy(); } catch (x) { /* ya cerrado */ } } return r; });
        const ex = leerExpediente(pags), dc = leerDocumento(pags);              // 2) expediente por regla fija
        if (!ex) { await aviso('No encontré el expediente (formato E + número / año) en <b>' + esc(f.name) + '</b>. ¿Es un oficio firmado?'); continue; }
        const ya = (datos().oficios || {})[ex.exp];                             // 3) repetidos: en la lista del titular o ya en el buzón
        if (ya) { await aviso(`<b>${esc(ex.exp)}</b> ya está en la lista: ${esc(tdoc(ya))} ${esc(ya.numero)}, estado <b>${esc(ya.sgd)}</b>, lleva ${plazoTxt(diasDe(ya))}, cargado por ${esc(ya.cargadoPor || R.titular)}.${ya.sgd === 'Respondido' ? '<br>Ya fue respondido; si volvió a la bandeja figurará como DEVUELTO.' : ''}<br><b>No se agregó.</b>`, 'Oficio repetido'); continue; }
        const buz = S.aportes.find(a => a.tipo === 'oficio' && a.exp === ex.exp && estAp(a.id) === 'nuevo');
        if (buz) { await aviso(`<b>${esc(ex.exp)}</b> ya lo dejó <b>${esc(buz.de)}</b> en el buzón (${esc(cuando(buz.hora))}). <b>No se agregó.</b>`, 'Oficio repetido'); continue; }
        const nombre = ex.exp.replace('/', '-') + '__' + L.usuario + '__' + stamp() + '.pdf';
        await FS.escribir(nombre, bytes, 'buzon');                              // 4) el PDF va al buzón; el titular lo lee con su propia clave
        await escribirAporte({ id: idAp(), tipo: 'oficio', exp: ex.exp, pdf: nombre, tipoDoc: dc.tipoDoc, numero: dc.numero, hora: new Date().toISOString() });
        await aviso(`<b>${esc(ex.exp)}</b> quedó en el buzón de <b>${esc(R.titular)}</b>.<br><br><b>Qué pasó:</b> tu PDF se copió a la carpeta compartida (buzon) y se anotó tu aporte. <b>Todavía no está en la tabla de ${esc(R.titular)}</b>.<br><b>Qué sigue:</b> el PNF de ${esc(R.titular)} (con el SGD abierto) revisa la carpeta al abrirse, al pulsar «Actualizar» y cada ${Math.round((R.seg || 60) / 60 * 10) / 10 || 1} min. Lo lee con su propia clave y lo agrega solo; si la lectura le deja dudas, se lo muestra para que lo revise. Al incorporarlo se borra el PDF del buzón y aquí, en «Mis aportes», pasará a «ya está en la tabla».<br>Tú no tienes que hacer nada más.`, 'Oficio dejado en el buzón');
      } catch (x) { await aviso('No se pudo agregar <b>' + esc(f.name) + '</b>: ' + esc(x.message || String(x))); }
      finally { libre(); }
    }
    S.leyendo = null; render();
  }
  async function procesarBuzon() {
    if (!esTit() || S.carpeta !== 'ok' || S.procesando) return;
    S.procesando = true; let cambios = false;
    try {
      D.aporteEst = D.aporteEst || {};
      S.buzonErr = {};
      // «enrevision» es un estado de versiones anteriores: la cola de revisión vive solo en memoria y se pierde al recargar el SGD, así que esos oficios se vuelven a leer
      for (const a of S.aportes.filter(x => x.tipo === 'oficio' && ['nuevo', 'enrevision'].includes(estAp(x.id)))) {
        if (D.oficios[a.exp]) { D.aporteEst[a.id] = 'incorporado'; cambios = true; continue; }          // el PDF NO se borra: solo se borra cuando PNF lo lee y lo agrega, o cuando tú lo descartas
        if (S.cola.some(c => c.aporteId === a.id)) continue;                    // ya está esperando tu revisión en esta sesión
        let bytes = null; try { bytes = await FS.leerBytes(a.pdf, 'buzon'); } catch (x) { S.buzonErr[a.id] = { txt: 'PNF no pudo abrir el archivo en la carpeta buzon.', det: errTxt(x) }; continue; }
        if (!bytes || !bytes.length) { if (Date.now() - new Date(a.hora).getTime() > 5 * 60000) S.buzonErr[a.id] = { txt: 'El archivo todavía no aparece en la carpeta buzon (puede que Drive aún lo esté sincronizando).', det: a.pdf }; continue; }
        let e; try { e = await leerArchivo(new File([bytes], a.pdf, { type: 'application/pdf' })); } catch (x) { S.buzonErr[a.id] = { txt: 'PNF no pudo leer el contenido del PDF.', det: errTxt(x) }; continue; }
        e.agregadoPor = a.de; e.aporteId = a.id;
        if (EXP_RE.test(e.exp) && TIPOS_DOC.includes(e.tipoDoc) && (e.cmp.ok || !e.ia)) { await incorporarOficio(e, a.de); D.aporteEst[a.id] = 'incorporado'; cambios = true; try { await FS.borrar(a.pdf, 'buzon'); } catch (x) { /* queda en el buzón */ } }
        else { S.cola.push(e); if (!S.lec) S.lec = { pend: e.id }; toast('⚠ El oficio de ' + a.de + ' (' + a.exp + ') necesita tu revisión.'); }   // sigue «nuevo» en el buzón hasta que lo agregues o descartes
      }
      if (cambios) await guardar('oficios recibidos de supervisores');
    } finally { S.procesando = false; }
  }


  /* ═══════════════ 15 · SGD: verificar la bandeja y enviar la respuesta (etapa 5) ═══════════════ */
  const quieto = ms => new Promise(r => setTimeout(r, ms));
  let enVuelo = 0;                                         // solicitudes AJAX del SGD en curso (el SGD usa RichFaces/A4J)
  (function () { try { const X = W.XMLHttpRequest, env = X.prototype.send; X.prototype.send = function () { enVuelo++; this.addEventListener('loadend', () => { enVuelo = Math.max(0, enVuelo - 1); }); return env.apply(this, arguments); }; } catch (e) { /* sin contador */ } })();
  const esperaVisible = () => qa('#panelEsperaContainer, #panelEspera').some(visible);   // «Por favor espere · Realizando operación…» del SGD
  async function esperarQuieto(maxMs) { const t0 = Date.now(); maxMs = maxMs || 15000; await quieto(250); while (Date.now() - t0 < maxMs) { if (enVuelo === 0 && !esperaVisible()) { await quieto(300); if (enVuelo === 0 && !esperaVisible()) return true; } await quieto(120); } return false; }
  async function esperar(fn, maxMs, paso) { const t0 = Date.now(); maxMs = maxMs || 12000; while (Date.now() - t0 < maxMs) { try { const v = fn(); if (v) return v; } catch (e) { /* sigue esperando */ } await quieto(paso || 150); } return null; }
  const visible = el => !!el && (el.offsetParent !== null || (el.getClientRects && el.getClientRects().length > 0));
  const norm = x => normNom(x);
  const qa = (css, raiz) => [...(raiz || document).querySelectorAll(css)];
  const q1 = (css, raiz) => (raiz || document).querySelector(css);
  const porId = (frag, tag, raiz) => qa((tag || '*') + '[id$="' + frag + '"]', raiz).find(visible) || qa((tag || '*') + '[id$="' + frag + '"]', raiz)[0] || null;
  const porIdContiene = (frag, tag, raiz) => qa((tag || '*') + '[id*="' + frag + '"]', raiz).find(visible) || null;
  function poner(el, v) { const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype, d = Object.getOwnPropertyDescriptor(proto, 'value'); if (d && d.set) d.set.call(el, v); else el.value = v; ['input', 'change', 'blur'].forEach(t => el.dispatchEvent(new Event(t, { bubbles: true }))); }
  function elegir(sel, texto) { const t = norm(texto), op = [...sel.options].find(o => norm(o.text) === t) || [...sel.options].find(o => norm(o.text).includes(t)); if (!op) return false; sel.value = op.value; sel.dispatchEvent(new Event('change', { bubbles: true })); return true; }
  const SGD_BANDEJA = () => location.origin + '/exedoc/pages/bandejaEntrada.seam';
  const avisoOtraPagina = () => paginaSGD() === 'bandeja' ? '' : '<div class="note warn" style="margin-top:8px">⚠ <b>Ahora estás en otra página del SGD.</b> PNF la cambiará a la bandeja de entrada: lo que tengas sin guardar en ella se perderá.</div>';
  const paginaSGD = () => /bandejaEntrada\.seam/.test(location.pathname) ? 'bandeja' : /verExpediente\.seam/.test(location.pathname) ? 'ver' : /crearGenericoPDF\.seam/.test(location.pathname) ? 'crear' : '';

  /* — leer la bandeja (en un marco oculto, sin molestar la página que estás viendo) — */
  function filasBandeja(doc) {
    const tb = doc.querySelector('tbody[id$="expedienteList:tb"]'); if (!tb) return null;
    return [...tb.querySelectorAll(':scope > tr')].map(tr => { const cs = [...tr.cells].map(td => td.innerText.replace(/\s+/g, ' ').trim()), i = cs.findIndex(c => /^E\d{1,9}\/\d{4}$/.test(c)); if (i < 0) return null; const f = cs.slice(i + 1).find(c => /^\d{2}\/\d{2}\/\d{4}/.test(c)) || ''; return { exp: cs[i], ingreso: f.slice(0, 10), hora: f.slice(11, 16), remitente: cs[i + 2] || '' }; }).filter(Boolean);
  }
  async function leerBandejaCompleta(doc) {
    const todas = [], vistos = new Set(); let pagina = 0;
    for (; pagina < 80; pagina++) {
      let f = null; await esperar(() => (f = filasBandeja(doc)), 15000, 250); if (!f) return pagina ? todas : null;
      f.forEach(x => { if (!vistos.has(x.exp)) { vistos.add(x.exp); todas.push(x); } });
      const act = doc.querySelector('.rich-datascr td.rich-datascr-act'), nActual = act ? +act.innerText.trim() : 0;
      const sig = nActual ? [...doc.querySelectorAll('.rich-datascr td.rich-datascr-inact')].find(td => +td.innerText.trim() === nActual + 1) : null;
      if (!sig) break;
      const firma = f.map(x => x.exp).join(',');
      sig.click();
      const cambio = await esperar(() => { const g = filasBandeja(doc); return g && g.map(x => x.exp).join(',') !== firma; }, 15000, 250); if (!cambio) break;
    }
    return todas;
  }
  async function bandejaEnMarco() {
    const fr = document.createElement('iframe'); fr.style.cssText = 'position:fixed;left:-9999px;top:0;width:1200px;height:800px;border:0;visibility:hidden'; fr.src = SGD_BANDEJA(); document.body.appendChild(fr);
    try { await new Promise((res, rej) => { fr.onload = res; fr.onerror = () => rej(new Error('no cargó la bandeja')); setTimeout(() => rej(new Error('la bandeja tardó demasiado en cargar')), 30000); }); return await leerBandejaCompleta(fr.contentDocument); }
    finally { fr.remove(); }
  }
  async function verificarBandeja() {
    if (!esTit()) return aviso('La verificación usa la bandeja del titular (' + esc(R.titular) + ').');
    if (S.verif) return; S.verif = true; toast('Verificando la bandeja del SGD (en segundo plano)…'); render();
    try {
      const filas = W.__PNF_BANDEJA ? await W.__PNF_BANDEJA() : await bandejaEnMarco();
      if (!filas || !filas.length) { await aviso('No pude leer la bandeja del SGD (¿sigue abierta tu sesión?). No se cambió nada.'); return; }
      const mapa = {}; filas.forEach(f => { mapa[f.exp] = f; });
      const hoy = hoyTxt(), res = { enBandeja: 0, devueltos: 0, respondidos: 0 };
      Object.keys(D.oficios).forEach(e => {
        const o = D.oficios[e], f = mapa[e];
        if (f) { res.enBandeja++; o.llegada = f.ingreso || o.llegada; if (o.sgd === 'Respondido') { o.sgd = 'Devuelto'; o.atendido = false; res.devueltos++; o.sgdTxt = 'volvió a la bandeja (había sido respondido el ' + (o.respondido || '—') + ')'; } else if (o.sgd !== 'Devuelto' && o.sgd !== 'Por despachar') { o.sgd = 'Pendiente'; o.sgdTxt = ''; } o.despachoClic = ''; }
        else if (o.sgd !== 'Respondido') { o.sgd = 'Respondido'; o.respondido = hoy; o.sgdTxt = 'ya no está en la bandeja (verificado el ' + hoy + ')'; res.respondidos++; }
        o.chk = new Date().toISOString();
      });
      D.chkAt = new Date().toISOString(); D.chkRes = res;
      await guardar('verificación de bandeja'); toast('Bandeja verificada: ' + res.enBandeja + ' en la bandeja' + (res.devueltos ? ' · ⚠ ' + res.devueltos + ' devuelto(s)' : '') + (res.respondidos ? ' · ' + res.respondidos + ' ya respondido(s)' : ''));
    } catch (e) { await aviso('No se pudo verificar la bandeja: ' + esc(e.message || String(e))); }
    finally { S.verif = false; render(); }
  }
  const devueltos = () => Object.keys(datos().oficios || {}).filter(e => datos().oficios[e].sgd === 'Devuelto' && !datos().oficios[e].atendido);

  /* — pantalla «Respuesta en SGD» — */
  const PASOS = ['Bandeja de entrada: filtrar por el expediente y abrirlo', 'Expediente: «Adjuntar respuesta» › Documento para firma (sólo PDF)', 'Formulario: tipo de documento, tipo de materia, materia, ciudad y actos con efecto sobre terceros', 'Adjuntar el PDF del memo (Agregar › archivo › Adjuntar)', 'Visador: Modificar › buscar usuario › Agregar › Volver', 'Firmante: Modificar › buscar usuario › Agregar › Volver', 'Distribución: Agregar › Mis Listas › lista guardada › Volver', 'Guardar', 'DESPACHAR — este clic es siempre tuyo'];
  const persona = u => datos().cfg.personas.find(p => p.usuario === u) || { usuario: u, nombre: u, cargo: '' };
  function vSgd() {
    const ex = exAct(), o = datos().oficios[ex];
    if (!o) return roBanner() + barraOficio(true) + frame('Respuesta en SGD', '<div class="vacio">Elige un oficio arriba o en «Oficios».</div>');
    const t = esTit(), c = conteo(o), lista = c.n > 0 && c.ok === c.n, C = datos().cfg, F = C.form, v = persona(C.visador), f = persona(C.firmante), dd = (k, x) => `<dt>${k}</dt><dd>${x}</dd>`;
    const aviso2 = o.sgd === 'Devuelto' ? `<div class="note mal" style="margin:8px 0 0">⚠ Volvió a tu bandeja después de responderlo (${esc(o.respondido || '—')}). PNF solo avisa, no bloquea.</div>` : o.sgd === 'Respondido' ? '<div class="note" style="margin:8px 0 0">Ya figura respondido. Si lo envías otra vez, será una respuesta adicional.</div>' : '';
    const lleva = o.sgd === 'Respondido' ? esc(o.sgdTxt || '') : 'llegó el ' + esc(o.llegada || o.fechaDoc || '—') + ' · lleva ' + plazoTxt(diasDe(o));
    return roBanner() + barraOficio(true) + '<div class="g2"><div>' +
      frame('Estado del expediente en el SGD', `<span class="estado ${claseEstado(o.sgd)}">${esc(o.sgd)}</span> ${lleva} <span style="font-size:11px">· verificado ${esc(cuando(o.chk))}</span>${aviso2}${o.sgd === 'Por despachar' ? `<div class="note warn" style="margin:8px 0 0">📤 Preparado en el SGD${o.preparado ? ' el ' + esc(cuando(o.preparado)) : ''}: solo falta pulsar «Despachar».${t ? `<div class="row" style="margin-top:6px">${btn('Ir a despachar', 'data-a="ir-despachar" data-id="' + esc(o.exp) + '"', 'def peq')}${btn('Crear el memo de nuevo', 'data-a="crear-nuevo" data-id="' + esc(o.exp) + '"', 'peq')}</div>` : ''}</div>` : ''}`) +
      frame('Memo a adjuntar', lista ? `<span class="estado e-v">memo listo</span><span class="pdfi" data-a="pdf-ver" title="Previsualizar el memo">PDF</span>${esc(nombreMemo(o))}.pdf <span style="font-size:11px;color:#555">(clic en el ícono para previsualizar · el PDF se genera justo antes de enviar)</span>` : `<span class="estado e-r">Falta confirmar ${c.n - c.ok} bloque(s) del análisis</span> ${btn('Ir al análisis', 'data-go="analisis"', 'peq')}`) +
      frame('Se rellenará así (valores de Configuración › SGD)', `<dl class="kv" style="grid-template-columns:170px minmax(0,1fr)">${dd('Tipo de respuesta', esc(F.tipoResp))}${dd('Tipo de documento', esc(F.tipoDoc))}${dd('Tipo de materia', esc(F.tipoMateria))}${dd('Materia', '<div style="max-height:34px;overflow:hidden" data-tip="' + esc(o.ant) + '">' + esc(o.ant) + '</div>')}${dd('Ciudad', esc(F.ciudad))}${dd('Actos con efecto sobre terceros', esc(F.actos))}${dd('Visador', esc(v.nombre) + ' (' + esc(v.usuario) + ')')}${dd('Firmante', esc(f.nombre) + ' (' + esc(f.usuario) + ')')}${dd('Distribución', 'Mis Listas › «' + esc(C.listaDist) + '»')}</dl>`) + '</div><div>' +
      frame('Enviar', `<div class="row">${btn('Enviar en SGD', 'data-a="sgd-enviar"' + (lista && t && !RUN ? '' : ' disabled'), 'def')}${btn('Verificar bandeja', 'data-a="verif"' + (t ? '' : ' disabled'), 'peq')}</div><div style="font-size:11.5px;margin-top:8px">Al pulsarlo, esta ventana se contrae para que veas el SGD y los pasos se van marcando en un recuadro abajo a la derecha. Al final «Despachar» queda resaltado con líneas segmentadas: lo pulsas tú y PNF detecta que el expediente quedó respondido. ${t ? '' : '<br><b>Solo el titular (' + esc(R.titular) + ') puede enviar al SGD.</b>'}</div>`) + '</div></div>' +
      `<div class="pie" style="display:flex;gap:8px;margin-top:10px"><button class="vb" data-go="memo">← Memo</button><span class="sp"></span><button class="vb" data-go="oficios">Volver a Oficios</button></div>`;
  }

  /* — la corrida de envío: una máquina de pasos que sobrevive a los cambios de página del SGD — */
  let RUN = GMX.get('pnf.run', null); if (RUN && Date.now() - RUN.ts > 40 * 60 * 1000) RUN = null;
  const saveRun = () => { if (RUN) { RUN.ts = Date.now(); GMX.set('pnf.run', RUN); } else GMX.set('pnf.run', null); };
  const PASOS_DESP = ['Bandeja de entrada: filtrar por el expediente y abrirlo', 'Expediente: comprobar que el memo está preparado', 'DESPACHAR — este clic es siempre tuyo'];
  function tarjeta() {
    const c = $('#card'); if (!RUN) { c.classList.remove('on'); return; }
    c.classList.add('on'); $('#cardT').textContent = 'PNF · ' + (RUN.modo === 'despachar' ? 'Despachando ' : 'Enviando ') + RUN.exp;
    $('#cardB').innerHTML = (RUN.modo === 'despachar' ? PASOS_DESP : PASOS).map((p, i) => `<div class="p ${i < RUN.paso ? 'ok' : ''}">${i < RUN.paso ? '☑' : i === RUN.paso ? (RUN.estado === 'error' ? '✖' : '▶') : '☐'} ${i + 1}. ${esc(p)}</div>`).join('') + `<div style="margin-top:6px;font-size:11.5px;${RUN.estado === 'error' ? 'color:#b00020' : ''}">${esc(RUN.msg || '')}</div><div class="row" style="margin-top:6px">${RUN.estado === 'espera' ? btn('Continuar', 'data-a="run-sigue"', 'def peq') : ''}${RUN.estado === 'error' && (RUN.fase !== 'crear' || RUN.sub === 0) ? btn('Reintentar este paso', 'data-a="run-reintenta"', 'def peq') : ''}${btn(RUN.estado === 'listo' ? 'Cerrar' : 'Cancelar', 'data-a="run-cancela"', 'peq')}</div>`;
  }
  function ocultarVentana(si) { if (si) { $('#win').classList.remove('on'); $('#fab').classList.remove('on'); } else { if (L.min) $('#fab').classList.add('on'); else $('#win').classList.add('on'); } }
  const msgRun = (m, estado) => { if (RUN) { RUN.msg = m; if (estado) RUN.estado = estado; saveRun(); tarjeta(); } };
  const falla = m => { msgRun(m, 'error'); throw new Error(m); };
  async function iniciarDespacho(ex, o) {
    RUN = { id: idAp(), exp: ex, modo: 'despachar', fase: 'bandeja', paso: 0, sub: 0, estado: 'corriendo', msg: 'Comenzando…', ts: Date.now(), ant: o.ant };
    saveRun(); ocultarVentana(true); tarjeta();
    if (paginaSGD() === 'bandeja') await correrRun(); else { msgRun('Abriendo la bandeja de entrada…'); location.href = SGD_BANDEJA(); }
  }
  async function enviarSGD(modo) {   // modo: 'ir' (solo despachar) | 'nuevo' (crear otro memo) | sin valor (pregunta si ya está preparado)
    const ex = exAct(), o = D.oficios[ex]; if (!o || !esTit() || RUN) return;
    if (modo === 'ir') return iniciarDespacho(ex, o);
    if (o.sgd === 'Por despachar' && !modo) {   // ya se creó el memo en el SGD: no se crea otro sin preguntar
      const r0 = await dialogo('Ya está preparado en el SGD', `<div><b>${esc(ex)}</b> ya quedó preparado en el SGD${o.preparado ? ' el ' + esc(cuando(o.preparado)) : ''}: solo falta pulsar <b>«Despachar»</b>.</div><div style="margin-top:6px">Si eliges «Crear de nuevo» se creará un segundo memo en el expediente.</div>${avisoOtraPagina()}`, [['ir', 'Ir a Despachar', true], ['nuevo', 'Crear de nuevo'], ['no', 'Cancelar']]);
      if (r0 !== 'ir' && r0 !== 'nuevo') return;
      if (r0 === 'ir') return iniciarDespacho(ex, o);
      modo = 'nuevo';
    }
    const c = conteo(o); if (!c.n || c.ok !== c.n) return aviso('Faltan bloques por confirmar en el análisis.');
    const C = D.cfg, v = persona(C.visador), f = persona(C.firmante);
    if (!C.listaDist || !v.usuario || !f.usuario) return aviso('Revisa en Configuración › SGD el visador, el firmante y la lista de distribución.');
    const r = await dialogo('Enviar en SGD', `<div>Se va a preparar la respuesta de <b>${esc(ex)}</b> en el SGD:</div><ol style="margin:6px 0 0;padding-left:20px;line-height:1.5"><li>abre el expediente desde la bandeja,</li><li>crea el «Documento para firma (sólo PDF)» y adjunta el memo,</li><li>pone visador <b>${esc(v.nombre)}</b>, firmante <b>${esc(f.nombre)}</b> y la lista «${esc(C.listaDist)}»,</li><li>pulsa Guardar${C.pausaGuardar ? ' (antes te deja revisar)' : ''}.</li></ol><div style="margin-top:8px"><b>«Despachar» lo pulsas tú.</b> Mientras trabaja no uses el SGD.</div>${modo === 'nuevo' ? '<div class="note warn" style="margin-top:8px">Ya hay un memo preparado en el SGD: se creará un <b>segundo</b>.</div>' : ''}${avisoOtraPagina()}`, [['ok', 'Enviar', true], ['no', 'Cancelar']]);
    if (r !== 'ok') return;
    toast('Generando el memo en PDF…'); let pdf;
    try { pdf = await memoPDF(ex); await PDFS.guardar('memo:' + ex, pdf.bytes); } catch (e) { return aviso('No se pudo generar el memo: ' + esc(e.message || String(e))); }
    RUN = { id: idAp(), exp: ex, fase: 'bandeja', paso: 0, sub: 0, estado: 'corriendo', msg: 'Comenzando…', ts: Date.now(), archivo: nombreMemo(o) + '.pdf', ant: o.ant, cfg: { tipoResp: C.form.tipoResp, tipoDoc: C.form.tipoDoc, tipoMateria: C.form.tipoMateria, ciudad: C.form.ciudad, actos: C.form.actos, visador: v, firmante: f, lista: C.listaDist, pausa: C.pausaGuardar } };
    saveRun(); ocultarVentana(true); tarjeta();
    if (paginaSGD() === 'bandeja') await correrRun(); else { msgRun('Abriendo la bandeja de entrada…'); location.href = SGD_BANDEJA(); }
  }
  async function reanudarRun() {
    if (!RUN) return; const pg = paginaSGD(); ocultarVentana(true); tarjeta();
    if (RUN.estado === 'listo') return;
    if (RUN.fase === 'bandeja' && pg !== 'bandeja') { msgRun('Abriendo la bandeja de entrada…'); location.href = SGD_BANDEJA(); return; }
    if (RUN.fase === 'crear' && pg === 'crear' && RUN.sub > 0) { msgRun('La página se recargó a mitad del formulario. Cancela y vuelve a enviar (o completa el resto a mano).', 'error'); return; }
    if ((RUN.fase === 'ver' && pg !== 'ver') || (RUN.fase === 'crear' && pg !== 'crear')) { await quieto(1500); if (paginaSGD() !== RUN.fase) { msgRun('Estoy en una página que no esperaba (' + (pg || 'otra') + '). Cancela y vuelve a intentar.', 'error'); return; } }
    await correrRun();
  }
  async function correrRun() {
    if (!RUN || RUN.estado === 'error' && !RUN.reintento) return; RUN.estado = 'corriendo'; RUN.reintento = false; saveRun();
    try {
      if (RUN.fase === 'bandeja') await faseBandeja(); else if (RUN.fase === 'ver') await faseVer(); else if (RUN.fase === 'crear') await faseCrear(); else if (RUN.fase === 'despachar') await faseDespachar();
    } catch (e) { if (RUN && RUN.estado !== 'error') msgRun(e.message || String(e), 'error'); }
  }
  // El expediente no está en la bandeja: se detiene y se pregunta qué es (despachado o error)
  async function expedienteNoEsta() {
    const ex = RUN.exp;
    msgRun('El expediente ' + ex + ' no aparece en tu bandeja.', 'error');
    const r = await dialogo('Expediente no encontrado', `<div><b>${esc(ex)}</b> no está en tu bandeja.</div>`, [['reint', 'Reintentar', true], ['desp', 'Ya despachado'], ['no', 'Cancelar']]);
    if (r === 'reint') { RUN.estado = 'corriendo'; RUN.paso = 0; saveRun(); tarjeta(); return faseBandeja(); }
    if (r === 'desp') {
      const o = D.oficios[ex]; if (o) { o.sgd = 'Respondido'; o.respondido = hoyTxt(); o.sgdTxt = 'marcado como despachado a mano el ' + hoyTxt() + ' ' + hh() + ' (no estaba en la bandeja)'; try { await guardar('expediente ' + ex + ' marcado como despachado'); } catch (x) { /* queda en este computador */ } }
      cancelarRun(true); toast('✔ ' + ex + ' quedó marcado como despachado.'); throw new Error('detenido: marcado como despachado');
    }
    cancelarRun(); throw new Error('detenido: cancelado');
  }
  // El memo preparado ya no está en el expediente (p. ej. lo eliminaron en el SGD): Reintentar · Crear de nuevo · Cancelar
  async function memoNoEsta() {
    const ex = RUN.exp;
    msgRun('No está el memo de ' + ex + ' en el SGD.', 'error');
    const r = await dialogo('Memo no encontrado', `<div>No está el memo de <b>${esc(ex)}</b> en el SGD.</div>`, [['reint', 'Reintentar', true], ['nuevo', 'Crear de nuevo'], ['no', 'Cancelar']]);
    if (r === 'reint') { RUN.fase = 'bandeja'; RUN.paso = 0; RUN.estado = 'corriendo'; saveRun(); tarjeta(); location.href = SGD_BANDEJA(); await quieto(30000); return; }
    if (r === 'nuevo') {
      const o = D.oficios[ex]; if (o) { o.sgd = 'Pendiente'; o.preparado = ''; o.sgdTxt = 'el memo preparado ya no estaba en el SGD (' + hoyTxt() + ' ' + hh() + ')'; try { await guardar('memo de ' + ex + ' ya no está en el SGD: vuelve a Pendiente'); } catch (x) { /* queda en este computador */ } }
      cancelarRun(true); S.exp = ex; S.sel = ex; S.vista = 'sgd'; render(); enviarSGD(); throw new Error('detenido: crear de nuevo');
    }
    cancelarRun(); throw new Error('detenido: cancelado');
  }
  async function faseBandeja() {
    msgRun('Buscando el expediente en la bandeja…');
    const filtro = await esperar(() => { const f = qa('table[id$="expedienteList"] input.rich-filter-input')[0]; return f && visible(f) ? f : null; }, 20000); if (!filtro) falla('No encontré el filtro de expedientes en la bandeja.');
    const num = RUN.exp.split('/')[0];
    poner(filtro, num);
    const ult = num.slice(-1); filtro.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, cancelable: true, key: ult, keyCode: ult.charCodeAt(0), which: ult.charCodeAt(0) }));   // el filtro del SGD se activa con «soltar una tecla» (onkeyup), no con «cambiar el texto»
    if (!(await esperarQuieto(20000))) falla('El SGD está lento: la búsqueda no terminó en 20 s. Cuando cargue, pulsa «Reintentar este paso».');
    // la lista ya está filtrada cuando todas sus filas son del número buscado; si no, todavía se ve la lista anterior
    const estado = () => { const reales = qa('tbody[id$="expedienteList:tb"] > tr').filter(tr => /E\d{1,9}\/\d{4}/.test(tr.innerText)); if (!reales.length) return { falta: true }; if (!reales.every(tr => tr.innerText.includes(num))) return null; const hay = reales.find(tr => tr.innerText.replace(/\s+/g, ' ').includes(RUN.exp)); return hay ? { fila: hay } : { falta: true }; };
    let r = await esperar(estado, 15000);
    if (r && r.falta) { await esperarQuieto(10000); await quieto(1500); r = estado(); }   // por si la lista aún se estaba cargando
    if (!r) falla('El SGD está lento: la lista no se filtró en 15 s. Cuando cargue, pulsa «Reintentar este paso».');
    if (r.falta) return expedienteNoEsta();
    let fila = r.fila; await quieto(400);
    if (!fila.isConnected) { const r2 = estado(); if (!r2 || !r2.fila) falla('La lista cambió mientras la leía. Pulsa «Reintentar este paso».'); fila = r2.fila; }
    const a = q1('a[id$=":verExpediente"]', fila); if (!a) falla('No encontré el enlace «Ver expediente» de la fila.');
    RUN.fase = 'ver'; RUN.paso = 1; saveRun(); tarjeta(); a.click();
    await quieto(20000);   // si la página no cambió, el clic no abrió el expediente
    RUN.fase = 'bandeja'; RUN.paso = 0; falla('No se abrió el expediente al pulsar «Ver» (¿el SGD está lento?). Pulsa «Reintentar este paso».');
  }
  async function faseVerDespachar() {
    msgRun('Comprobando que el memo está preparado en el expediente…');
    const filas = await esperar(() => { const r = qa('tr.rich-subtable-row').filter(tr => tr.children.length >= 10); return r.length ? r : null; }, 20000); if (!filas) falla('No veo la lista de documentos del expediente.');
    if (!document.body.innerText.includes(RUN.exp)) falla('El expediente abierto no es ' + RUN.exp + '. Cancelo para no equivocarme.');
    const ses = leerSesionSGD(), ant = norm(RUN.ant || '').slice(0, 60);
    const mios = filas.filter(tr => { const t = norm(tr.innerText); return t.includes('MEMORANDUM') && t.includes('CREADO') && (!ses || mismoNombre(ses.nombre, tr.innerText)) && (ant.length < 10 || t.includes(ant)); });
    if (!mios.length) return memoNoEsta();
    if (mios.length > 1) falla('Hay ' + mios.length + ' memos iguales en estado CREADO en este expediente (¿se envió dos veces?). Elimina el que sobre en el SGD y vuelve a intentar.');
    RUN.fase = 'despachar'; RUN.paso = 2; saveRun(); tarjeta(); await faseDespachar();
  }
  async function faseVer() {
    if (RUN.modo === 'despachar') return faseVerDespachar();
    msgRun('Abriendo «Adjuntar respuesta»…');
    const boton = await esperar(() => q1('#panelDocumentoOriginal_header a.buttonright'), 20000); if (!boton) falla('No encontré el botón «Adjuntar respuesta» en el expediente.');
    if (!document.body.innerText.includes(RUN.exp)) falla('El expediente abierto no es ' + RUN.exp + '. Cancelo para no equivocarme.');
    boton.click();
    const sel = await esperar(() => { const s = q1('#panelCrearRespuestaContainer select'); return s && visible(s) ? s : null; }, 10000); if (!sel) falla('No se abrió la ventana «Adjuntar respuesta».');
    RUN.fase = 'crear'; RUN.paso = 2; RUN.sub = 0; saveRun(); tarjeta();
    if (!elegir(sel, RUN.cfg.tipoResp)) falla('No encontré la opción «' + RUN.cfg.tipoResp + '».');
  }
  async function pasoForm(nombre, fn) { msgRun(nombre + '…'); const r = await fn(); await esperarQuieto(15000); return r; }
  async function faseCrear() {
    const C = RUN.cfg, form = await esperar(() => porId('decorateCiudad:inputCiudad', 'input'), 25000); if (!form) falla('No cargó el formulario del documento.');
    RUN.paso = 2; RUN.sub = 1; saveRun(); tarjeta();
    await pasoForm('Tipo de documento «' + C.tipoDoc + '»', async () => { const s = porIdContiene('idDocumentosElectronicosList', 'select'); if (!s || !elegir(s, C.tipoDoc)) falla('No pude elegir el tipo de documento «' + C.tipoDoc + '».'); });
    await pasoForm('Tipo de materia «' + C.tipoMateria + '»', async () => { const s = porId('comboTipoMateria', 'select'); if (!s || !elegir(s, C.tipoMateria)) falla('No pude elegir el tipo de materia «' + C.tipoMateria + '».'); });
    await pasoForm('Materia (el ANT. del oficio)', async () => { const t = porId('decorateMateria:ioMateria', 'textarea'); if (!t) falla('No encontré el campo Materia.'); poner(t, RUN.ant); });
    await pasoForm('Ciudad «' + C.ciudad + '»', async () => { const i = porId('decorateCiudad:inputCiudad', 'input'); if (!i) falla('No encontré el campo Ciudad.'); poner(i, C.ciudad); });
    await pasoForm('Actos con efecto sobre terceros «' + C.actos + '»', async () => { const s = qa('select', q1('form[id="crearGenerico"]') || document).find(x => visible(x) && [...x.options].some(o => norm(o.text) === 'SI') && [...x.options].some(o => norm(o.text) === 'NO') && x.options.length <= 3); if (!s || !elegir(s, C.actos)) falla('No pude elegir «Actos con efecto sobre terceros».'); });
    RUN.paso = 3; RUN.sub = 2; saveRun();
    await adjuntarMemo(); RUN.paso = 4; RUN.sub = 3; saveRun();
    await elegirPersona('Visador', 'botonAgregarVisador', 'buscarUsuarioCrearResolucionVisacion', 'botonVolverAgregarVisadores', C.visador); RUN.paso = 5; RUN.sub = 4; saveRun();
    await elegirPersona('Firmante', 'botonAgregarFirmante', 'buscarUsuarioCrearResolucionFirma', 'botonVolverAgregarFirmantes', C.firmante); RUN.paso = 6; RUN.sub = 5; saveRun();
    await distribucion(C.lista); RUN.paso = 7; RUN.sub = 6; saveRun();
    verificarFormulario(C);
    if (C.pausa) { msgRun('Revisa el formulario. Cuando esté bien, pulsa «Continuar» para guardar.', 'espera'); return; }
    await guardarFormulario();
  }
  function verificarFormulario(C) {
    const faltan = [], txt = document.body.innerText;
    const td = porIdContiene('idDocumentosElectronicosList', 'select'); if (!td || norm(td.options[td.selectedIndex].text) !== norm(C.tipoDoc)) faltan.push('tipo de documento');
    const mt = porId('decorateMateria:ioMateria', 'textarea'); if (!mt || !mt.value.trim()) faltan.push('materia');
    if (!norm(txt).includes(norm(RUN.archivo.replace(/\.pdf$/i, '')))) faltan.push('el PDF adjunto');
    if (!norm(txt).includes(norm(C.visador.nombre).split(' ').slice(0, 3).join(' '))) faltan.push('el visador');
    if (!norm(txt).includes(norm(C.firmante.nombre).split(' ').slice(0, 3).join(' '))) faltan.push('el firmante');
    if (faltan.length) falla('Antes de guardar no veo en el formulario: ' + faltan.join(', ') + '. No guardé nada.');
  }
  async function guardarFormulario() {
    const b = porId('crearGenerico:guardar', 'input') || qa('input[type=button],input[type=submit]').find(x => x.value === 'Guardar' && visible(x)); if (!b) falla('No encontré el botón Guardar.');
    msgRun('Guardando…'); RUN.fase = 'despachar'; RUN.paso = 8; RUN.sub = 7; saveRun(); tarjeta(); b.click();
  }
  async function adjuntarMemo() {
    msgRun('Adjuntando el PDF del memo…');
    const ag = porId('crearGenerico:botonAgregarAntecedentes', 'input') || porId('botonAgregarAntecedentes', 'input'); if (!ag) falla('No encontré el botón «Agregar» de los documentos adjuntos.');
    ag.click(); const fi = await esperar(() => qa('input[type=file][id*="uploadFile"]').find(visible) || qa('input[type=file][id*="uploadFile"]')[0], 12000); if (!fi) falla('No se abrió la ventana para adjuntar el archivo.');
    const bytes = await PDFS.leer('memo:' + RUN.exp); if (!bytes) falla('No encontré el memo generado (¿se limpió el almacenamiento?). Vuelve a enviar.');
    const dt = new DataTransfer(); dt.items.add(new File([bytes], RUN.archivo, { type: 'application/pdf' })); fi.files = dt.files; fi.dispatchEvent(new Event('change', { bubbles: true })); await esperarQuieto(20000);
    const bt = await esperar(() => porId('botonAgregarArchivo', 'input'), 10000); if (!bt) falla('No encontré el botón «Adjuntar».');
    bt.click(); await esperarQuieto(20000);
    if (!(await esperar(() => norm(document.body.innerText).includes(norm(RUN.archivo.replace(/\.pdf$/i, ''))), 10000))) falla('No veo el archivo adjunto en el formulario.');
  }
  async function elegirPersona(rotulo, idBoton, idBusca, idVolver, p) {
    msgRun(rotulo + ': ' + p.nombre + '…');
    const mod = porId(idBoton, 'input'); if (!mod) falla('No encontré el botón «Modificar» del ' + rotulo.toLowerCase() + '.');
    mod.click(); const campo = await esperar(() => porId(idBusca + ':idUsuario', 'input'), 12000); if (!campo || !visible(campo)) falla('No se abrió la búsqueda de usuarios del ' + rotulo.toLowerCase() + '.');
    // El SGD a veces deja escritos datos «fantasma» (Nombre, Apellido) de una búsqueda anterior y la búsqueda combina todos los campos: se limpian antes de buscar.
    const cont = campo.closest('[id$="panelBuscarUsuarioFiltro_body"]') || campo.closest('form') || campo.closest('table');
    const otros = () => qa('input[type=text],input:not([type]),textarea', cont).filter(i => i !== campo && visible(i));
    otros().forEach(i => { if (i.value) poner(i, ''); });
    if (otros().some(i => i.value)) falla('No pude limpiar los campos Nombre y Apellido de la búsqueda de usuarios. Límpialos a mano y pulsa «Reintentar este paso».');
    poner(campo, p.usuario); const bus = porId(idBusca + ':addDestinatarioExterno', 'input'); if (!bus) falla('No encontré el botón «Buscar».'); bus.click(); await esperarQuieto(15000);
    // En el SGD real el botón «Agregar» vive dentro de su propia tablita (…:N:agregar) dentro de la fila: la fila con los nombres es la primera fila superior que tenga texto además del botón.
    const filaDeBoton = i => { let tr = i.closest('tr'); while (tr && norm(tr.innerText).replace(/\bAGREGAR\b/g, '').trim().length < 3) tr = tr.parentElement && tr.parentElement.closest('tr'); return tr; };
    const filas = await esperar(() => { const fs = qa('input[id*="' + idBusca + '"][id$=":botonAgregarDestinatario"]').filter(visible).map(filaDeBoton).filter(Boolean); return fs.length ? fs : null; }, 12000);
    if (!filas) falla('La búsqueda de «' + p.usuario + '» no devolvió resultados.');
    const cand = filas.filter(tr => mismoNombre(p.nombre, tr.innerText) || norm(tr.innerText).includes(norm(p.nombre).split(' ').slice(0, 3).join(' ')));
    if (cand.length !== 1) falla(cand.length ? 'La búsqueda de «' + p.usuario + '» devolvió ' + cand.length + ' personas con ese nombre; no elijo a ciegas.' : 'Entre los resultados no está «' + p.nombre + '» (nunca elijo por posición).');
    q1('input[id$=":botonAgregarDestinatario"]', cand[0]).click(); await esperarQuieto(15000);
    const vol = await esperar(() => porId(idVolver, 'input'), 10000); if (!vol) falla('No encontré el botón «Volver».'); vol.click(); await esperarQuieto(20000);
  }
  async function distribucion(lista) {
    msgRun('Distribución: lista «' + lista + '»…');
    const ag = porId('crearGenerico:botonAgregarDistribucion', 'input') || porId('botonAgregarDistribucion', 'input'); if (!ag) falla('No encontré el botón «Agregar» de la distribución.');
    ag.click(); const pest = await esperar(() => { const td = qa('td[id$="tabListaDestinatarios_lbl"]').find(visible); return td; }, 12000); if (!pest) falla('No se abrió la ventana de distribución.');
    pest.click(); await esperarQuieto(15000);
    const filas = await esperar(() => { const fs = qa('a[id$=":botonAgregarListaDestFrecuentes"]').filter(visible).map(i => i.closest('tr')); return fs.length ? fs : null; }, 12000); if (!filas) falla('No veo ninguna lista en «Mis Listas». ¿Existe «' + lista + '»?');
    const cand = filas.filter(tr => norm(tr.innerText).includes(norm(lista))); if (cand.length !== 1) falla(cand.length ? 'Hay ' + cand.length + ' listas que coinciden con «' + lista + '».' : 'No existe la lista «' + lista + '» en tus «Mis Listas». Créala en el SGD y reintenta.');
    q1('a[id$=":botonAgregarListaDestFrecuentes"]', cand[0]).click(); await esperarQuieto(15000);
    const vol = await esperar(() => porId('botonVolverAgregarDistribucion', 'input'), 10000); if (!vol) falla('No encontré el botón «Volver».'); vol.click(); await esperarQuieto(20000);
  }
  async function faseDespachar() {
    msgRun('Buscando el botón «Despachar»…');
    const b = await esperar(() => qa('input[type=button],input[type=submit],a.button,a.buttonright,button').find(x => /^Despachar$/i.test((x.value || x.innerText || '').trim()) && visible(x)), 25000);
    if (!b) falla('No encontré el botón «Despachar» (el expediente pudo quedar guardado). Búscalo tú en el expediente.');
    const o0 = D.oficios[RUN.exp];   // el memo ya está guardado en el SGD: queda anotado que solo falta Despachar
    if (o0 && (RUN.modo !== 'despachar' || o0.sgd !== 'Por despachar')) { o0.sgd = 'Por despachar'; o0.preparado = new Date().toISOString(); o0.sgdTxt = 'preparado en el SGD el ' + hoyTxt() + ' ' + hh() + '; falta Despachar'; try { await guardar('expediente ' + RUN.exp + ' preparado en el SGD: falta Despachar'); } catch (x) { /* queda en este computador */ } }
    RUN.paso = RUN.modo === 'despachar' ? 2 : 8; saveRun(); b.scrollIntoView({ behavior: 'smooth', block: 'center' }); resaltar(b);
    msgRun('Listo para despachar: pulsa tú el botón «Despachar» (resaltado con líneas segmentadas).', 'espera'); RUN.estado = 'esperaDespacho'; saveRun(); tarjeta();
    const h = async ev => { const t = ev.target.closest && ev.target.closest('input,a,button'); if (!t || !/^Despachar$/i.test((t.value || t.innerText || '').trim())) return; document.removeEventListener('click', h, true); await marcarDespachado(); };
    document.addEventListener('click', h, true);
  }
  function resaltar(el) {
    const st = document.createElement('style'); st.id = 'pnf-hl'; st.textContent = '@keyframes pnfdash { from { outline-color:#e8590c; } to { outline-color:#ffc107; } } .pnf-despachar { outline:3px dashed #e8590c !important; outline-offset:4px !important; animation:pnfdash .7s infinite alternate !important; }'; document.head.appendChild(st); el.classList.add('pnf-despachar');
  }
  async function marcarDespachado() {
    const ex = RUN && RUN.exp; RUN.estado = 'listo'; RUN.paso = RUN.modo === 'despachar' ? 3 : 9; RUN.msg = '✔ Despachado. El expediente figura como respondido; se confirmará al verificar la bandeja.'; saveRun(); tarjeta();
    try { const o = D.oficios[ex]; if (o) { o.sgd = 'Respondido'; o.respondido = hoyTxt(); o.sgdTxt = 'despachado el ' + hoyTxt() + ' ' + hh(); o.despachoClic = new Date().toISOString(); o.memoOk = true; await guardar('expediente ' + ex + ' respondido'); } } catch (e) { /* se confirma al verificar */ }
  }
  function cancelarRun(cierra) { const ya = RUN && RUN.estado === 'listo'; RUN = null; saveRun(); tarjeta(); const h = document.getElementById('pnf-hl'); if (h) h.remove(); ocultarVentana(false); if (!cierra && !ya) toast('Envío cancelado'); render(); }

  /* ═══════════════ 8 · Arranque y sondeo de avisos ═══════════════ */
  let timer = null;
  function iniciarSondeo() {
    clearInterval(timer);
    timer = setInterval(async () => { if (document.visibilityState !== 'visible') return; if (verificarSesion() === 'cambio') await identificar(); await sincronizar(); const a = root.activeElement; if (a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName)) pintarEstado(); else render(); }, Math.max(10, R.seg || 60) * 1000);
  }
  document.addEventListener('visibilitychange', async () => { if (document.visibilityState === 'visible') { if (verificarSesion() === 'cambio') await identificar(); await sincronizar(); render(); } });

  async function iniciar() {
    if (L.min) { $('#fab').classList.add('on'); } else { $('#win').classList.add('on'); }
    render();
    try { await conectar(false); } catch (e) { /* sin carpeta */ }
    verificarSesion();
    if (!L.usuario) { if (L.min) restaurar(); await identificar(); }
    await sincronizar(); if (!esTit() && L.usuario && S.vista === 'oficios') S.vista = 'sup'; render();
    if (RUN) { reanudarRun(); iniciarSondeo(); return; }
    if (!L.bienvenida && S.carpeta === 'sin') { L.bienvenida = true; saveL(); if (L.min) restaurar(); await guiaCarpeta(); render(); }
    if (esTit() && S.ofrecer) { await adoptar(); render(); }
    iniciarSondeo();
  }
  if (W.__PNF_TEST) W.__PNF = { S, L, get R() { return R; }, set R(v) { R = v; }, get D() { return D; }, set D(v) { D = v; }, get V() { return V; }, FS, root, render, guardar, guardarRoles, sincronizar, conectar, elegirCarpeta, verificarBandeja, enviarSGD, filasBandeja, leerBandejaCompleta, correrRun, reanudarRun, cancelarRun, paginaSGD, get RUN() { return RUN; }, set RUN(v) { RUN = v; }, aplicarReglas, sugerir, conteo, textoAn, hojaMemo, paginasMemo, memoPDF, memoWord, nombreMemo, filasExcel, exportarExcel, leerAportes, escribirAporte, procesarBuzon, agregarOficioSupervisor, incorporarOficio, descargarMemo, accionAnalisis, vAnalisis, vMemo, vExcel, vSup, leerArchivo, cargarArchivos, agregarPend, descartarPend, errTxt, comparar, agrupar, filasOf, renderTabla, leerExpediente, leerDocumento, leerEncabezado, filasNormal, lineasPagina, lector, PDFS, cbAbrir, quitarOficio, corregirDatos, entradaPend, leerSesionSGD, verificarSesion, mismoNombre, candidatosPorNombre, guiaCarpeta, alertas, esTit, esAdm, pinOk, cmpVer, aplicarTraspaso, cargarDeCarpeta, copiaSeguridad, escribirDatos, iniciarSondeo, minimizar, restaurar, actualizar, identificar, VERSION };
  iniciar();
})();
