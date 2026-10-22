/* Hub Comex — app de Ventas para registrar despachos y traslados con fotos y generar su informe en Word.
 *
 * - Ingreso: Microsoft (MSAL) o modo prueba (se elige Ventas / Soporte). La sesión queda guardada en el dispositivo.
 * - Fotos: OneDrive → /<carpetaRaiz>/<Despachos|Traslados>/<AAAA-MM-DD - nombre>/Foto_001.jpg
 * - Cola de subida: IndexedDB. Si no hay señal, las fotos esperan en el celular.
 * - Registro: /<carpetaRaiz>/_registro-hub-comex.json (despachos y traslados, cada uno con su tipo).
 * - Datos e inspección: se guardan primero en el dispositivo y se sincronizan cuando hay señal.
 * - Informe: informe.js arma el .docx con la librería "docx" (se carga solo al descargar el primer informe).
 */
(function () {
  'use strict';

  const CFG = window.HUB_CONFIG || {};
  const SCOPES = CFG.permisos || ['User.Read', 'Files.ReadWrite'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const RAIZ = CFG.carpetaRaiz || 'Hub Comex';
  const REGISTRO = `/me/drive/root:/${encodeURIComponent(RAIZ)}/_registro-hub-comex.json`;
  const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js';
  const LIMITE_SIMPLE = 4 * 1024 * 1024; // Graph acepta PUT directo hasta 4 MB
  const ES_FOTO = /\.(jpe?g|png|heic|heif|webp)$/i;
  // Modo prueba (config modo: 'prueba', o …/#demo): sin Microsoft. Se elige Ventas o Soporte al ingresar
  // y los registros y fotos se guardan solo en este dispositivo.
  const DEMO = CFG.modo !== 'microsoft' || location.hash === '#demo' || window.HUB_DEMO === true;
  const DEMO_CUENTAS = {
    ventas: { username: 'ventas-sales@healthyfood.com.ec', name: 'Ventas' },
    soporte: { username: 'notificaciones@healthyfood.com.ec', name: 'Soporte' },
    otro: { username: 'otro.usuario@healthyfood.com.ec', name: 'Otro usuario' }
  };
  const SESION_DEMO = 'hc.demo.sesion';

  const TIPOS = {
    despacho: { plural: 'Despachos', uno: 'despacho', Uno: 'Despacho', carpeta: 'Despachos', ej: 'Ej. Pulpa mango Goya contenedor C7' },
    traslado: { plural: 'Traslados', uno: 'traslado', Uno: 'Traslado', carpeta: 'Traslados', ej: 'Ej. Pulpa maracuyá a bodega Guayaquil' }
  };
  const T = d => TIPOS[(d && d.tipo) || d] || TIPOS.despacho;
  const CAMPOS = [
    ['cliente', 'Cliente', 'Ej. Goya Foods'],
    ['destino', 'Destino', 'Ej. Jersey City, EE. UU.'],
    ['lugarCarga', 'Lugar de carga', 'Ej. Planta Quevedo'],
    ['pedido', 'Pedido / OC', 'Ej. OC-2026-0418']
  ];
  const INSPECCION = (window.HCInforme && window.HCInforme.INSPECCION) || [
    ['exterior', 'Exterior, puertas y estructura'], ['interior', 'Interior limpio, seco y sin olores'],
    ['piso', 'Piso, paredes y techo íntegros'], ['refrigeracion', 'Refrigeración operativa'],
    ['llantas', 'Llantas, luces y elementos de seguridad'], ['apta', 'Unidad apta para carga de alimentos']
  ];

  // ---------- Utilidades ----------
  const ICONS = {
    plus: '<path d="M12 5v14M5 12h14"/>',
    camera: '<path d="M4 8h3l2-3h6l2 3h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="13.5" r="3.6"/>',
    image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="1.8"/><path d="m21 16-5-5-9 9"/>',
    folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    ext: '<path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.2-4.2"/>',
    home: '<path d="M4 11 12 4l8 7v8a1 1 0 0 1-1 1h-5v-6h-4v6H5a1 1 0 0 1-1-1z"/>',
    truck: '<path d="M2 6h11v10H2zM13 9h4.5l3.5 3.5V16h-8"/><circle cx="6" cy="17.5" r="1.8"/><circle cx="17" cy="17.5" r="1.8"/>',
    swap: '<path d="M4 8h13l-3.5-3.5M20 16H7l3.5 3.5"/>',
    doc: '<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4M9 12h6M9 16h6"/>',
    back: '<path d="M15 18l-6-6 6-6"/>',
    chev: '<path d="m9 6 6 6-6 6"/>',
    check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
    x: '<path d="M6 6l12 12M18 6 6 18"/>',
    cloud: '<path d="M7 18a4.5 4.5 0 0 1-.6-8.96A6 6 0 0 1 18 9.5a4.25 4.25 0 0 1-.5 8.5z"/><path d="m9.5 13.5 2.5-2.5 2.5 2.5M12 11v5"/>',
    cloudok: '<path d="M7 18a4.5 4.5 0 0 1-.6-8.96A6 6 0 0 1 18 9.5a4.25 4.25 0 0 1-.5 8.5z"/><path d="m9.3 13.6 2 2 3.6-3.8"/>',
    nosignal: '<path d="M3 3l18 18M8.5 16.5a5 5 0 0 1 7 0M5 13a10 10 0 0 1 4-2.4M12 20h.01M19 13a10 10 0 0 0-2.3-1.7M2 9.3A15 15 0 0 1 6.6 6.5M22 9.3A15 15 0 0 0 11 5.5"/>',
    clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
    alert: '<path d="M12 4 2.5 20h19z"/><path d="M12 10v4.5M12 17.5h.01"/>',
    lock: '<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
    redo: '<path d="M4 12a8 8 0 1 0 2.4-5.7L4 8.5"/><path d="M4 4v4.5h4.5"/>',
    info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5M12 8h.01"/>',
    logout: '<path d="M15 4h3a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-3M10 16l-4-4 4-4M6 12h10"/>',
    spin: '<path d="M12 3a9 9 0 1 0 9 9"/>'
  };
  const ic = n => `<svg class="i" viewBox="0 0 24 24" aria-hidden="true">${ICONS[n]}</svg>`;
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pad = n => String(n).padStart(2, '0');
  const pad3 = n => String(n).padStart(3, '0');
  const isoDay = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fmtDate = d => new Date(d).toLocaleDateString('es-EC', { day: 'numeric', month: 'short', year: 'numeric' });
  const fmtTime = d => new Date(d).toLocaleTimeString('es-EC', { hour: '2-digit', minute: '2-digit' });
  const fmtSize = b => b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.round(b / 1024) + ' KB';
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

  // OneDrive no acepta  \ / : * ? " < > |  (ni # % por compatibilidad) ni espacios o puntos al final.
  const PROHIBIDOS = /[\\/:*?"<>|#%]/g;
  function nombreCarpeta(raw, fecha) {
    const quitados = [...new Set(raw.match(PROHIBIDOS) || [])];
    const limpio = raw.replace(PROHIBIDOS, '-').replace(/\s+/g, ' ').trim().replace(/[ .]+$/, '').slice(0, 100);
    return { nombre: limpio ? `${isoDay(fecha)} - ${limpio}` : '', quitados };
  }

  class SinSenal extends Error {}

  // ---------- Almacenamiento local ----------
  const idb = (() => {
    let abierto;
    const abrir = () => abierto || (abierto = new Promise((res, rej) => {
      const r = indexedDB.open('hub-comex', 1);
      r.onupgradeneeded = () => { r.result.createObjectStore('cola', { keyPath: 'id' }); r.result.createObjectStore('kv'); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }));
    const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
    const store = async (name, mode) => (await abrir()).transaction(name, mode).objectStore(name);
    return {
      cola: async () => req((await store('cola')).getAll()),
      poner: async item => req((await store('cola', 'readwrite')).put(item)),
      quitar: async id => req((await store('cola', 'readwrite')).delete(id)),
      leer: async k => req((await store('kv')).get(k)),
      guardar: async (k, v) => req((await store('kv', 'readwrite')).put(v, k))
    };
  })();

  const cache = {
    leer(k) { try { return JSON.parse(localStorage.getItem('hc.' + k)); } catch (_) { return null; } },
    guardar(k, v) { try { localStorage.setItem('hc.' + k, JSON.stringify(v)); } catch (_) {} },
    quitar(k) { try { localStorage.removeItem('hc.' + k); } catch (_) {} }
  };

  // ---------- Sesión Microsoft ----------
  let pca, cuenta;
  async function iniciarSesionMsal() {
    pca = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: 'https://login.microsoftonline.com/' + CFG.tenantId,
        redirectUri: location.origin + location.pathname,
        navigateToLoginRequestUrl: false
      },
      cache: { cacheLocation: 'localStorage' }
    });
    await pca.initialize();
    const r = await pca.handleRedirectPromise();
    cuenta = (r && r.account) || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (cuenta) pca.setActiveAccount(cuenta);
  }
  const ingresar = () => pca.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
  const salir = () => pca.logoutRedirect({ account: cuenta, postLogoutRedirectUri: location.origin + location.pathname });

  async function token() {
    if (DEMO) return 'demo';
    try {
      return (await pca.acquireTokenSilent({ scopes: SCOPES, account: cuenta })).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        await pca.acquireTokenRedirect({ scopes: SCOPES, account: cuenta });
        return new Promise(() => {}); // la página se va al login de Microsoft
      }
      if (!navigator.onLine || /network|timeout|fetch/i.test(`${e.errorCode} ${e.message}`)) throw new SinSenal();
      throw e;
    }
  }

  async function graph(ruta, op = {}) {
    if (DEMO) return demoGraph(ruta, op);
    const t = await token();
    const headers = { Authorization: 'Bearer ' + t, ...(op.headers || {}) };
    let body = op.body;
    if (op.json) { headers['Content-Type'] = 'application/json'; body = JSON.stringify(op.json); }
    let r;
    try { r = await fetch(GRAPH + ruta, { method: op.method || 'GET', headers, body }); }
    catch (_) { throw new SinSenal(); }
    if (!r.ok) {
      const err = new Error(`Graph respondió ${r.status}`);
      err.status = r.status;
      try { err.detalle = (await r.json()).error; } catch (_) {}
      throw err;
    }
    if (r.status === 204) return null;
    return op.raw ? r : r.json();
  }

  // ---------- OneDrive simulado (solo modo prueba) ----------
  const demo = { carpetas: new Map(), registro: null, n: 0 };
  function demoError(status) { const e = new Error('Demo ' + status); e.status = status; return e; }
  function demoCarpeta(nombre, padre) {
    const id = 'demo-' + uid() + (++demo.n);
    const hermanos = [...demo.carpetas.values()].filter(c => c.padre === padre).map(c => c.name);
    let name = nombre, k = 1;
    while (hermanos.includes(name)) name = `${nombre} ${k++}`;
    demo.carpetas.set(id, { id, name, padre, fotos: [] });
    return { id, name, webUrl: '#' };
  }
  function demoBuscarRuta(ruta) {
    let padre = null, c;
    for (const parte of ruta.split('/').filter(Boolean)) {
      c = [...demo.carpetas.values()].find(x => x.padre === padre && x.name === parte);
      if (!c) return null;
      padre = c.id;
    }
    return c;
  }
  function demoSemilla() {
    const raiz = demoCarpeta(RAIZ, null).id;
    const sub = { despacho: demoCarpeta('Despachos', raiz).id, traslado: demoCarpeta('Traslados', raiz).id };
    const si = { v: 'si' };
    const ejemplos = [
      ['despacho', 'Pulpa mango Goya contenedor C6', 3, 38, 'Goya Foods Inc.', 'Jersey City, EE. UU.'],
      ['traslado', 'Pulpa maracuyá a bodega Guayaquil', 8, 12, '', 'Bodega Guayaquil']
    ];
    const lista = ejemplos.map(([tipo, nombre, dias, fotos, cliente, destino]) => {
      const fecha = new Date(Date.now() - dias * 864e5);
      const fin = new Date(fecha.getTime() + 3 * 36e5);
      const c = demoCarpeta(nombreCarpeta(nombre, fecha).nombre, sub[tipo]);
      demo.carpetas.get(c.id).fotos = Array.from({ length: fotos }, (_, i) => ({ id: c.id + '-' + i, name: `Foto_${pad3(i + 1)}.jpg` }));
      return {
        id: c.id, tipo, nombre, carpeta: c.name, webUrl: '#', fecha: fecha.toISOString(), estado: 'cerrado', cerrado: fin.toISOString(), fotos,
        cliente, destino, lugarCarga: 'Planta Quevedo', pedido: '',
        inspeccion: { exterior: si, interior: si, piso: si, refrigeracion: si, llantas: si, apta: si }
      };
    });
    demo.registro = JSON.stringify({ despachos: lista });
  }
  // Lo simulado se guarda en este dispositivo: datos en localStorage, fotos en IndexedDB.
  const demoClave = () => 'demo.datos.v3.' + cuenta.username;
  function demoGuardar() {
    cache.guardar(demoClave(), {
      n: demo.n, registro: demo.registro,
      carpetas: [...demo.carpetas.values()].map(c => ({ id: c.id, name: c.name, padre: c.padre, fotos: c.fotos.map(f => ({ id: f.id, name: f.name })) }))
    });
  }
  async function demoCargar() {
    const d = cache.leer(demoClave());
    if (!d) { demo.carpetas = new Map(); demoSemilla(); demoGuardar(); return; }
    Object.assign(demo, { n: d.n, registro: d.registro, carpetas: new Map() });
    for (const c of d.carpetas) {
      for (const f of c.fotos) {
        const blob = await idb.leer('demofoto.' + f.id).catch(() => null);
        if (blob) f.url = URL.createObjectURL(blob);
      }
      demo.carpetas.set(c.id, c);
    }
  }
  async function demoGraph(ruta, op) {
    await new Promise(r => setTimeout(r, 200 + Math.random() * 400));
    if (!navigator.onLine) throw new SinSenal();
    const metodo = op.method || 'GET';
    let m;
    if (ruta.startsWith(REGISTRO)) {
      if (metodo === 'PUT') { demo.registro = op.body; demoGuardar(); return {}; }
      if (!demo.registro) throw demoError(404);
      return { '@microsoft.graph.downloadUrl': 'data:application/json,' + encodeURIComponent(demo.registro) };
    }
    if ((m = /^\/me\/drive\/root:\/(.+)$/.exec(ruta))) {
      const c = demoBuscarRuta(decodeURIComponent(m[1]));
      if (!c) throw demoError(404);
      return { id: c.id };
    }
    if (ruta === '/me/drive/root/children' && metodo === 'POST') { const c = demoCarpeta(op.json.name, null); demoGuardar(); return c; }
    if ((m = /^\/me\/drive\/items\/([^/:?]+)\/children$/.exec(ruta)) && metodo === 'POST') {
      const c = demoCarpeta(op.json.name, m[1]);
      demoGuardar();
      return c;
    }
    if ((m = /^\/me\/drive\/items\/([^/:?]+):\/([^:]+):\/content/.exec(ruta))) {
      const c = demo.carpetas.get(m[1]);
      if (!c) throw demoError(404);
      const f = { id: c.id + '-' + c.fotos.length + '-' + Date.now(), name: decodeURIComponent(m[2]), url: URL.createObjectURL(op.body) };
      await idb.guardar('demofoto.' + f.id, op.body);
      c.fotos.push(f);
      demoGuardar();
      return { id: f.id, name: f.name };
    }
    if ((m = /^\/me\/drive\/items\/([^/:?]+)\/children\?/.exec(ruta))) {
      const c = demo.carpetas.get(m[1]);
      if (!c) throw demoError(404);
      return { value: c.fotos.map(f => ({ id: f.id, name: f.name, file: {}, thumbnails: f.url ? [{ medium: { url: f.url } }] : [] })) };
    }
    if ((m = /^\/me\/drive\/items\/([^/:?]+)$/.exec(ruta)) && metodo === 'DELETE') {
      for (const c of demo.carpetas.values()) {
        const i = c.fotos.findIndex(f => f.id === m[1]);
        if (i < 0) continue;
        c.fotos.splice(i, 1);
        await idb.guardar('demofoto.' + m[1], null);
        demoGuardar();
        return null;
      }
      throw demoError(404);
    }
    if ((m = /^\/me\/drive\/items\/([^/:?]+)\?/.exec(ruta))) {
      const c = demo.carpetas.get(m[1]);
      if (!c) throw demoError(404);
      return { id: c.id, folder: { childCount: c.fotos.length } };
    }
    throw demoError(400);
  }

  // ---------- Registro (despachos y traslados) ----------
  let registro = { despachos: [] };
  const clavePend = () => 'pendientes.' + cuenta.username;

  // Une cambios de datos/inspección sobre un registro (la inspección se mezcla punto por punto).
  function mezclar(obj, cambio) {
    for (const k of Object.keys(cambio)) {
      if (k === 'inspeccion') {
        obj.inspeccion = obj.inspeccion || {};
        for (const p of Object.keys(cambio.inspeccion)) obj.inspeccion[p] = { ...(obj.inspeccion[p] || {}), ...cambio.inspeccion[p] };
      } else obj[k] = cambio[k];
    }
    return obj;
  }
  function aplicarPendientes(data) {
    const pend = cache.leer(clavePend()) || {};
    for (const id of Object.keys(pend)) { const x = data.despachos.find(y => y.id === id); if (x) mezclar(x, pend[id]); }
    return data;
  }

  async function leerRegistro() {
    let data;
    try {
      const meta = await graph(REGISTRO);
      let r;
      try { r = await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' }); }
      catch (_) { throw new SinSenal(); }
      if (!r.ok) { const e = new Error('No se pudo leer el registro'); e.status = r.status; throw e; }
      data = JSON.parse(await r.text());
    } catch (e) {
      if (e.status !== 404) throw e;
      data = { despachos: [] };
    }
    if (!Array.isArray(data.despachos)) data.despachos = [];
    data.despachos.forEach(x => { if (!x.tipo) x.tipo = 'despacho'; });
    return data;
  }
  async function cargarRegistro() {
    registro = aplicarPendientes(await leerRegistro());
    cache.guardar('registro.' + cuenta.username, registro);
    return registro;
  }
  // Lee lo último de OneDrive, aplica los cambios pendientes y el cambio pedido, y lo guarda.
  async function cambiarRegistro(cambio) {
    const pendAntes = JSON.stringify(cache.leer(clavePend()) || {});
    const data = aplicarPendientes(await leerRegistro());
    cambio(data);
    await graph(REGISTRO + ':/content', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data, null, 1) });
    if (JSON.stringify(cache.leer(clavePend()) || {}) === pendAntes) cache.quitar(clavePend());
    registro = aplicarPendientes(data);
    cache.guardar('registro.' + cuenta.username, registro);
  }
  const buscar = id => registro.despachos.find(d => d.id === id);

  // Datos e inspección: se guardan al instante en el dispositivo y se suben cuando hay señal.
  function guardarDetalles(id, cambio) {
    const d = buscar(id);
    if (!d) return;
    mezclar(d, cambio);
    const pend = cache.leer(clavePend()) || {};
    pend[id] = mezclar(pend[id] || {}, JSON.parse(JSON.stringify(cambio)));
    cache.guardar(clavePend(), pend);
    cache.guardar('registro.' + cuenta.username, registro);
    sincronizar();
  }
  let sincronizando = false, otraVez = false;
  async function sincronizar() {
    if (!cuenta || !Object.keys(cache.leer(clavePend()) || {}).length) return;
    if (sincronizando) { otraVez = true; return; }
    sincronizando = true;
    try { await cambiarRegistro(() => {}); } catch (_) { /* sin señal: se reintenta luego */ }
    finally {
      sincronizando = false;
      if (otraVez) { otraVez = false; sincronizar(); }
    }
  }

  const idCarpetas = {};
  async function asegurarCarpeta(ruta, padreId) {
    if (idCarpetas[ruta]) return idCarpetas[ruta];
    try {
      idCarpetas[ruta] = (await graph(`/me/drive/root:/${ruta.split('/').map(encodeURIComponent).join('/')}`)).id;
    } catch (e) {
      if (e.status !== 404) throw e;
      const destino = padreId ? `/me/drive/items/${padreId}/children` : '/me/drive/root/children';
      idCarpetas[ruta] = (await graph(destino, { method: 'POST', json: { name: ruta.split('/').pop(), folder: {}, '@microsoft.graph.conflictBehavior': 'fail' } })).id;
    }
    return idCarpetas[ruta];
  }
  async function carpetaTipo(tipo) {
    const raiz = await asegurarCarpeta(RAIZ);
    return asegurarCarpeta(`${RAIZ}/${T(tipo).carpeta}`, raiz);
  }

  async function crearRegistro(tipo, datos) {
    const fecha = new Date();
    const { nombre } = nombreCarpeta(datos.nombre, fecha);
    const padre = await carpetaTipo(tipo);
    // "rename": si ya existe una carpeta con ese nombre, OneDrive agrega un número al final
    const f = await graph(`/me/drive/items/${padre}/children`, { method: 'POST', json: { name: nombre, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' } });
    const d = {
      id: f.id, tipo, nombre: datos.nombre.trim().replace(/\s+/g, ' '), carpeta: f.name, webUrl: f.webUrl,
      fecha: fecha.toISOString(), estado: 'proceso', cerrado: null, fotos: 0, inspeccion: {}
    };
    CAMPOS.forEach(([k]) => { d[k] = (datos[k] || '').trim(); });
    await cambiarRegistro(r => r.despachos.unshift(d));
    return d;
  }

  async function contarFotos(id) {
    const r = await graph(`/me/drive/items/${id}/children?$select=id,name,file&$top=999`);
    return r.value.filter(f => f.file && ES_FOTO.test(f.name)).length;
  }

  // ---------- Cola de subida ----------
  let cola = [];               // fotos pendientes (de esta cuenta)
  let subiendoId = null;       // foto que se está subiendo ahora
  let sinSenal = !navigator.onLine;
  let trabajando = false, reintento;
  const urlsCola = new Map();  // id de la foto → URL local para mostrarla
  const recientes = new Map(); // nombre del archivo subido → URL local (mientras OneDrive genera la miniatura)

  async function refrescarCola() {
    cola = (await idb.cola()).filter(x => x.cuenta === cuenta.username).sort((a, b) => a.creado - b.creado);
    for (const it of cola) if (!urlsCola.has(it.id)) urlsCola.set(it.id, URL.createObjectURL(it.blob));
  }
  const pendientesDe = id => cola.filter(x => x.despachoId === id);

  async function subirArchivo(it) {
    const ruta = `/me/drive/items/${it.folderId}:/${encodeURIComponent(it.archivo)}:`;
    if (DEMO || it.blob.size <= LIMITE_SIMPLE) {
      return graph(`${ruta}/content?@microsoft.graph.conflictBehavior=rename`, { method: 'PUT', headers: { 'Content-Type': it.blob.type || 'image/jpeg' }, body: it.blob });
    }
    const s = await graph(`${ruta}/createUploadSession`, { method: 'POST', json: { item: { '@microsoft.graph.conflictBehavior': 'rename' } } });
    const PARTE = 320 * 1024 * 10; // múltiplo de 320 KB, como pide Graph
    let ultimo;
    for (let i = 0; i < it.blob.size; i += PARTE) {
      const trozo = it.blob.slice(i, i + PARTE);
      let r;
      try { r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${i}-${i + trozo.size - 1}/${it.blob.size}` }, body: trozo }); }
      catch (_) { throw new SinSenal(); }
      if (!r.ok) { const e = new Error('Subida fallida'); e.status = r.status; throw e; }
      ultimo = r;
    }
    return ultimo.json();
  }

  // Las fotos aceptadas quedan en el celular ("sin subir") hasta que se toca "Subir fotos";
  // recién ahí pasan a la cola (listo: true) y se suben, esperando señal si hace falta.
  const porSubir = id => pendientesDe(id).filter(x => !x.listo);
  const enCola = () => cola.filter(x => x.listo && !x.error);
  async function subirFotos(d) {
    const lote = porSubir(d.id);
    for (const it of lote) { it.listo = true; await idb.poner(it); }
    await refrescarCola();
    procesarCola();
    return lote.length;
  }

  async function procesarCola() {
    if (trabajando || !cuenta) return;
    trabajando = true;
    clearTimeout(reintento);
    sincronizar();
    try {
      await refrescarCola();
      for (;;) {
        const it = cola.find(x => x.listo && !x.error);
        if (!it) break;
        subiendoId = it.id; refrescar();
        try {
          const subido = await subirArchivo(it);
          recientes.set(subido.name, urlsCola.get(it.id));
          urlsCola.delete(it.id);
          await idb.quitar(it.id);
          sinSenal = false;
          if (vista.id === it.despachoId) vista.nube.push({ id: subido.id, name: subido.name });
          const x = buscar(it.despachoId);
          if (x) guardarDetalles(x.id, { fotos: (x.fotos || 0) + 1 });
        } catch (e) {
          if (e instanceof SinSenal) { sinSenal = true; reintento = setTimeout(procesarCola, 20000); break; }
          if (e.status === 404) { it.error = 'La carpeta ya no existe en OneDrive.'; await idb.poner(it); }
          else { reintento = setTimeout(procesarCola, 15000); break; } // OneDrive ocupado (429/5xx): reintenta luego
        } finally {
          subiendoId = null;
          await refrescarCola();
          refrescar();
        }
      }
    } finally {
      trabajando = false;
      refrescar();
    }
  }
  window.addEventListener('online', () => { sinSenal = false; refrescar(); procesarCola(); });
  window.addEventListener('offline', () => { sinSenal = true; refrescar(); });
  document.addEventListener('visibilitychange', () => { if (!document.hidden) procesarCola(); });

  async function encolar(items, d) {
    let n = Math.max((await idb.leer('contador.' + d.id)) || 0, vista.id === d.id ? vista.nube.length + pendientesDe(d.id).length : 0);
    for (const it of items) {
      n++;
      await idb.poner({
        id: uid(), cuenta: cuenta.username, despachoId: d.id, folderId: d.id, despacho: d.nombre,
        archivo: `Foto_${pad3(n)}.${it.ext}`, blob: it.blob, creado: Date.now(), listo: false
      });
    }
    await idb.guardar('contador.' + d.id, n);
    await refrescarCola();
    return n;
  }

  // Reduce la foto a 1920 px por lado y JPEG 80 %: buena para informes y rápida de subir.
  function comprimir(file) {
    return new Promise(resolve => {
      const src = URL.createObjectURL(file);
      const original = () => resolve({ blob: file, url: src, ext: (file.name.split('.').pop() || 'jpg').toLowerCase(), orig: file.size });
      const img = new Image();
      img.onload = () => {
        const k = Math.min(1, 1920 / Math.max(img.width, img.height));
        const c = document.createElement('canvas');
        c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        c.toBlob(b => {
          if (!b || b.size >= file.size) return original();
          URL.revokeObjectURL(src);
          resolve({ blob: b, url: URL.createObjectURL(b), ext: 'jpg', orig: file.size });
        }, 'image/jpeg', 0.8);
      };
      img.onerror = original;
      img.src = src;
    });
  }

  // ---------- Informe en Word ----------
  function cargarLibreriaWord() {
    if (window.docx && window.docx.Packer) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = DOCX_URL;
      s.onload = () => (window.docx && window.docx.Packer ? res() : rej(new Error('Librería docx no disponible')));
      s.onerror = () => { s.remove(); rej(new SinSenal()); };
      document.head.appendChild(s);
    });
  }
  async function descargarInforme(d) {
    if (window.HUB_DEMO) { toast('En esta vista de prueba no se pueden descargar archivos. Usa la app publicada en GitHub.'); return; }
    st.generando = d.id; render();
    try {
      await cargarLibreriaWord();
      const doc = window.HCInforme.crear(window.docx, d, CFG.informe || {});
      const blob = await window.docx.Packer.toBlob(doc);
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = window.HCInforme.nombreArchivo(d);
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 10000);
      toast('Informe descargado. Ábrelo en Word para completarlo.');
    } catch (e) {
      console.error(e);
      toast(e instanceof SinSenal ? 'Sin señal: la primera vez se necesita internet para preparar el informe.' : 'No se pudo generar el informe. Avisa a soporte.');
    }
    st.generando = null; render();
  }

  // ---------- Estado de la pantalla ----------
  // Pantallas: 'areas' (elegir Despachos/Traslados) → dentro de cada área: 'modulo' (inicio del área),
  // 'nuevo', 'registros' (buscar y descargar informes) y 'despacho' (un registro abierto).
  const app = document.getElementById('app');
  const cam = document.getElementById('cam');
  const gal = document.getElementById('gal');
  const st = {
    view: 'boot', tipo: 'despacho', prev: 'modulo', rol: null, menu: false, revision: null, ocupado: false,
    nuevo: {}, q: '', desde: '', hasta: '', estado: 'todos', error: null, abiertos: {}, borrar: null, generando: null
  };
  let vista = { id: null, nube: [], cargando: false, error: null };
  const EN_AREA = ['modulo', 'nuevo', 'registros', 'despacho'];

  function toast(msg) {
    document.querySelector('.toast')?.remove();
    const t = document.createElement('div');
    t.className = 'toast'; t.setAttribute('role', 'status'); t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 3800);
  }
  function errorAMensaje(e, accion) {
    if (e instanceof SinSenal) return `No hay señal. Necesitas conexión para ${accion}.`;
    return `No se pudo ${accion}. Intenta de nuevo; si sigue fallando, avisa a soporte.`;
  }
  function ir(view, extra) {
    Object.assign(st, extra || {});
    st.view = view; st.menu = false;
    render(); window.scrollTo(0, 0);
  }
  // Re-pinta solo pantallas que muestran el avance, y nunca mientras se escribe en un campo.
  let repintar = false;
  function refrescar() {
    if (st.revision || !['despacho', 'areas', 'modulo'].includes(st.view)) return;
    const a = document.activeElement;
    if (a && app.contains(a) && /INPUT|TEXTAREA/.test(a.tagName)) { repintar = true; return; }
    render();
  }
  document.addEventListener('focusout', () => {
    if (!repintar) return;
    setTimeout(() => {
      const a = document.activeElement;
      if (a && app.contains(a) && /INPUT|TEXTAREA/.test(a.tagName)) return;
      repintar = false; refrescar();
    }, 250);
  });
  const saludo = () => { const h = new Date().getHours(); return h < 12 ? 'Buenos días' : h < 19 ? 'Buenas tardes' : 'Buenas noches'; };
  const ordenados = tipo => registro.despachos.filter(d => !tipo || d.tipo === tipo).sort((a, b) => b.fecha.localeCompare(a.fecha));
  const inspHechos = d => INSPECCION.filter(([k]) => d.inspeccion && d.inspeccion[k] && d.inspeccion[k].v).length;
  const datosHechos = d => CAMPOS.filter(([k]) => (d[k] || '').trim()).length;
  const delMes = d => { const f = new Date(d.fecha), h = new Date(); return f.getMonth() === h.getMonth() && f.getFullYear() === h.getFullYear(); };

  // ---------- Pantallas ----------
  function vIngreso() {
    return `<main class="wrap"><section class="login">
      <div class="mark">HC</div>
      <h1>Hub Comex</h1>
      <button class="btn primary big block" data-act="ingresar">${ic('lock')}Iniciar sesión</button>
      ${DEMO && st.opcionesDemo ? `<div class="demo-box">
        <span class="eyebrow">Elige con qué cuenta entrar</span>
        <button class="btn" data-act="demologin" data-u="ventas">Ventas<span>uso normal</span></button>
        <button class="btn" data-act="demologin" data-u="soporte">Soporte<span>pruebas</span></button>
        <button class="btn" data-act="demologin" data-u="otro">Otro correo<span>sin acceso</span></button>
      </div>` : ''}
    </section></main>`;
  }
  function vSinAcceso() {
    return `<main class="wrap"><section class="login">
      <div class="mark bad">${ic('lock')}</div>
      <div><h1>Esta app es solo para Ventas</h1>
      <p class="muted" style="margin-top:8px">La cuenta <code>${esc(cuenta.username)}</code> no tiene acceso. Si necesitas entrar, contacta al Área de Innovación y Mantenimiento.</p></div>
      <button class="btn block" data-act="salir">Usar otra cuenta</button>
    </section></main>`;
  }
  function vFalla() {
    return `<main class="wrap"><section class="login">
      <div class="mark bad">${ic('alert')}</div>
      <div><h1>No se pudo abrir la app</h1><p class="muted" style="margin-top:8px">${esc(st.error)}</p></div>
      <button class="btn block" data-act="recargar">${ic('redo')}Intentar de nuevo</button>
    </section></main>`;
  }

  function encabezado(ancho) {
    const nombre = cuenta.name || cuenta.username;
    const enArea = EN_AREA.includes(st.view);
    return `<header class="top ${ancho ? 'wide' : ''}"><div class="top-in">
      ${enArea
        ? `<button class="areas-btn" data-act="areas" aria-label="Volver a elegir área">${ic('back')}<span>Áreas</span></button>
           <div class="top-name">${T(st.tipo).plural}<small>Hub Comex</small></div>`
        : `<div class="mark">HC</div><div class="top-name">Hub Comex<small>Área de Ventas</small></div>`}
      <button class="who" data-act="menu" aria-label="Cuenta"><span class="ava">${esc(nombre.trim()[0] || 'V').toUpperCase()}</span><span>${esc(nombre.split(' ')[0])}</span></button>
    </div></header>`;
  }
  function barra() {
    if (!EN_AREA.includes(st.view)) return '';
    const v = st.view;
    const t = T(st.tipo);
    return `<nav class="nav" aria-label="${t.plural}"><div class="nav-in">
      <button class="nav-b" data-act="ir" data-v="modulo" ${v === 'modulo' ? 'aria-current="page"' : ''}>${ic('home')}Inicio</button>
      <button class="nav-plus" data-act="nuevo" aria-label="Nuevo ${t.uno}">${ic('plus')}</button>
      <button class="nav-b" data-act="registros" ${v === 'registros' ? 'aria-current="page"' : ''}>${ic('list')}Registros</button>
    </div></nav>`;
  }
  function tarjeta(d) {
    const sinSubir = porSubir(d.id).length;
    const sub = [fmtDate(d.fecha), d.cliente ? esc(d.cliente) : '', `${d.fotos} fotos${sinSubir ? ` · ${sinSubir} sin subir` : ''}`].filter(Boolean).join(' · ');
    return `<div class="card">
      <button class="main" data-act="abrir" data-id="${esc(d.id)}">
        <span class="name">${esc(d.nombre)}</span>
        <span class="meta">${sub} <span class="pill ${d.estado}">${d.estado === 'proceso' ? 'En proceso' : 'Cerrado'}</span></span>
      </button>
      <button class="icon-btn" data-act="informe" data-id="${esc(d.id)}" aria-label="Descargar informe en Word" title="Informe Word">${st.generando === d.id ? ic('spin') : ic('doc')}</button>
      <a class="icon-btn" href="${esc(d.webUrl)}" target="_blank" rel="noopener" aria-label="Abrir carpeta en OneDrive" title="Abrir carpeta">${ic('folder')}</a>
    </div>`;
  }
  function notas() {
    const pend = enCola().length;
    return `${DEMO ? `<div class="note info">${ic('info')}<span><b>Versión de prueba.</b> Los registros y fotos se guardan solo en este dispositivo; todavía no se suben a OneDrive.</span></div>` : ''}
      ${st.rol === 'soporte' ? `<div class="note info">${ic('info')}<span>Estás en modo soporte. Tus pruebas se guardan en tu propio OneDrive, no en el de Ventas.</span></div>` : ''}
      ${pend ? `<div class="note ${sinSenal ? 'bad' : ''}">${ic(sinSenal ? 'nosignal' : 'cloud')}<span>${pend} ${pend === 1 ? 'foto' : 'fotos'} ${sinSenal ? 'esperando señal para subir.' : 'subiendo…'}</span></div>` : ''}`;
  }

  // Inicio: elegir el área, como en el Hub de Recepción.
  function vAreas() {
    const nombre = st.rol === 'soporte' ? 'Soporte' : 'Ventas';
    const area = (tipo, icono, desc) => {
      const lista = ordenados(tipo);
      const abiertos = lista.filter(d => d.estado === 'proceso').length;
      return `<button class="area" data-act="area" data-t="${tipo}">
        <span class="area-ic">${ic(icono)}</span>
        <b>${T(tipo).plural}</b>
        <small>${desc}</small>
        <span class="area-n">${lista.filter(delMes).length} este mes${abiertos ? ` · <em>${abiertos} en proceso</em>` : ''}</span>
      </button>`;
    };
    return `<main class="wrap areas-wrap">
      <div class="areas-head"><h1>${saludo()}, ${nombre}</h1><p>Elige el área</p></div>
      ${notas()}
      <div class="areas">
        ${area('despacho', 'truck', 'Exportación de producto · fotos e informe')}
        ${area('traslado', 'swap', 'Traslado de producto · fotos e informe')}
      </div>
    </main>`;
  }

  // Inicio del área: crear uno nuevo y ver los que siguen abiertos.
  function vModulo() {
    const t = T(st.tipo);
    const lista = ordenados(st.tipo);
    const abiertos = lista.filter(d => d.estado === 'proceso');
    const recientes = lista.filter(d => d.estado !== 'proceso').slice(0, 3);
    return `<main class="wrap">
      <div class="greet"><h1>${t.plural}</h1>
        <p>${new Date().toLocaleDateString('es-EC', { weekday: 'long', day: 'numeric', month: 'long' })} · ${lista.filter(delMes).length} este mes</p></div>
      ${notas()}
      <button class="cta" data-act="nuevo"><span class="plus">${ic('plus')}</span><span><b>Nuevo ${t.uno}</b><small>Datos básicos y a tomar fotos</small></span></button>
      ${abiertos.length ? `<section class="list"><span class="eyebrow">En proceso</span>${abiertos.map(tarjeta).join('')}</section>` : ''}
      <section class="list">
        <div class="sec-h"><h2>Últimos cerrados</h2>${lista.length ? `<button class="link" data-act="registros">Ver registros</button>` : ''}</div>
        ${recientes.length ? recientes.map(tarjeta).join('')
          : `<div class="empty">${lista.length ? `Los ${t.plural.toLowerCase()} cerrados aparecerán aquí.` : `Aún no hay ${t.plural.toLowerCase()}. Toca <b>Nuevo ${t.uno}</b> para crear el primero.`}</div>`}
      </section>
    </main>`;
  }

  function filtrados() {
    const q = st.q.trim().toLowerCase();
    return ordenados(st.tipo).filter(d => {
      if (q && ![d.nombre, d.carpeta, d.cliente, d.destino, d.pedido].some(x => (x || '').toLowerCase().includes(q))) return false;
      const dia = isoDay(new Date(d.fecha));
      if (st.desde && dia < st.desde) return false;
      if (st.hasta && dia > st.hasta) return false;
      if (st.estado !== 'todos' && d.estado !== st.estado) return false;
      return true;
    });
  }
  function listaRegistro() {
    const filas = filtrados();
    const t = T(st.tipo);
    const hayFiltro = st.q || st.desde || st.hasta || st.estado !== 'todos';
    return `<div class="sec-h"><span class="count">${filas.length} ${filas.length === 1 ? t.uno : t.plural.toLowerCase()}</span>
        ${hayFiltro ? `<button class="link" data-act="limpiar">Quitar filtros</button>` : ''}</div>
      ${filas.length ? filas.map(tarjeta).join('')
        : `<div class="empty">${ordenados(st.tipo).length ? `Ningún ${t.uno} coincide con la búsqueda.` : `Aún no hay ${t.plural.toLowerCase()} registrados.`}</div>`}`;
  }
  // Registros: buscar, filtrar y descargar el informe del que se necesite.
  function vRegistros() {
    const t = T(st.tipo);
    const chip = (v, l) => `<button class="chip" data-act="estado" data-e="${v}" aria-pressed="${st.estado === v}">${l}</button>`;
    return `<main class="wrap wide">
      <div><h1 style="font-size:28px">Registros de ${t.plural.toLowerCase()}</h1>
        <p class="muted" style="margin-top:6px">Toca ${ic('doc')} para descargar el informe en Word o ${ic('folder')} para abrir la carpeta de fotos.</p></div>
      <div class="reg-layout">
        <div class="filters">
          <div class="search">${ic('search')}<input class="input" id="q" type="search" placeholder="Buscar por nombre, cliente o pedido" value="${esc(st.q)}" autocomplete="off"></div>
          <div class="dates">
            <label for="desde">Desde<input class="input" type="date" id="desde" value="${st.desde}"></label>
            <label for="hasta">Hasta<input class="input" type="date" id="hasta" value="${st.hasta}"></label>
          </div>
          <div class="chips">${chip('todos', 'Todos')}${chip('proceso', 'En proceso')}${chip('cerrado', 'Cerrados')}</div>
        </div>
        <div class="list" id="lista">${listaRegistro()}</div>
      </div>
    </main>`;
  }

  // Listas para elegir: el catálogo de config.js (cuando se cargue) más lo ya escrito en registros anteriores.
  function sugerencias() {
    const catalogo = CFG.catalogos || {};
    return CAMPOS.map(([k]) => {
      const vals = [...new Set([...(catalogo[k] || []), ...registro.despachos.map(d => (d[k] || '').trim())].filter(Boolean))].slice(0, 200);
      return `<datalist id="dl-${k}">${vals.map(v => `<option value="${esc(v)}"></option>`).join('')}</datalist>`;
    }).join('');
  }
  function campo(k, etiqueta, ej, valor, prefijo) {
    return `<div class="field"><label for="${prefijo}-${k}">${etiqueta}</label>
      <input class="input" id="${prefijo}-${k}" data-campo="${k}" list="dl-${k}" autocomplete="off" placeholder="${esc(ej)}" value="${esc(valor || '')}" maxlength="120"></div>`;
  }
  function vistaRuta(final, quitados) {
    return `<span class="eyebrow">Se creará la carpeta</span>
      <div class="crumbs">${ic('folder')}OneDrive <span>›</span> ${esc(RAIZ)} <span>›</span> ${esc(T(st.tipo).carpeta)} <span>›</span> <b>${final ? esc(final) : '…'}</b></div>
      ${quitados.length ? `<p class="warn">${ic('alert')}<span>OneDrive no acepta ${quitados.map(c => `<code class="mono">${esc(c)}</code>`).join(' ')} en nombres de carpeta; se cambiaron por guiones.</span></p>` : ''}`;
  }
  function vNuevo() {
    const t = T(st.tipo);
    const n = st.nuevo;
    const { nombre, quitados } = nombreCarpeta(n.nombre || '', new Date());
    return `<main class="wrap">
      <h1 style="font-size:28px">Nuevo ${t.uno}</h1>
      <div class="field">
        <label for="nuevo-nombre">Nombre del ${t.uno} <span class="req">obligatorio</span></label>
        <input class="input" id="nuevo-nombre" data-campo="nombre" autocomplete="off" placeholder="${esc(t.ej)}" value="${esc(n.nombre || '')}" maxlength="120">
        <p class="help">Será el nombre de la carpeta. Escríbelo como lo vas a buscar después.</p>
      </div>
      <div class="preview-path" id="ruta">${vistaRuta(nombre, quitados)}</div>
      <div class="form-grid">${CAMPOS.map(([k, e, ej]) => campo(k, e, ej, n[k], 'nuevo')).join('')}</div>
      <p class="help">Al crear se guarda la fecha y la hora de inicio; la hora de fin se guarda al cerrar. Estos datos se pueden completar o corregir después.</p>
      ${sugerencias()}
      <button class="btn primary big block" data-act="crear" id="crear" ${(n.nombre || '').trim().length >= 3 && !st.ocupado ? '' : 'disabled'}>${st.ocupado ? ic('spin') + 'Creando carpeta…' : ic('camera') + 'Crear y tomar fotos'}</button>
    </main>`;
  }

  function cajaSubida(d) {
    const pend = pendientesDe(d.id);
    const borradores = pend.filter(x => !x.listo).length;
    const conError = pend.filter(x => x.error);
    const esperando = pend.filter(x => x.listo && !x.error).length;
    const subidas = vista.nube.length;
    const total = subidas + esperando;
    const pct = total ? Math.round(subidas / total * 100) : 0;
    let cls = '', icon = 'cloud', linea, ayuda;
    if (vista.cargando && !total) { linea = 'Revisando fotos…'; ayuda = ''; }
    else if (!total) { linea = 'Aún no hay fotos subidas'; ayuda = ''; }
    else if (!esperando) { cls = 'ok'; icon = 'cloudok'; linea = `${total} ${total === 1 ? 'foto subida' : 'fotos subidas'}`; ayuda = DEMO ? 'Están guardadas en la carpeta.' : 'Ya están en OneDrive.'; }
    else if (sinSenal) { cls = 'off'; icon = 'nosignal'; linea = `${subidas} de ${total} fotos subidas`; ayuda = `<b>Sin señal.</b> ${esperando} ${esperando === 1 ? 'foto espera' : 'fotos esperan'} en el celular y se ${esperando === 1 ? 'subirá' : 'subirán'} solas cuando vuelva la conexión.`; }
    else { cls = 'going'; icon = 'cloud'; linea = `${subidas} de ${total} fotos subidas`; ayuda = 'Subiendo en segundo plano. No cierres la app hasta que terminen.'; }
    return `${borradores ? `<div class="stage">
        <div class="stage-tx"><b>${borradores} ${borradores === 1 ? 'foto sin subir' : 'fotos sin subir'}</b><span>Revisa y elimina las que no sirvan. Cuando estén todas, súbelas.</span></div>
        <button class="btn primary" data-act="subirfotos">${ic('cloud')}Subir ${borradores}</button>
      </div>` : ''}
      ${total || vista.cargando ? `<div class="upbox ${cls}">
        <div class="upline">${ic(icon)}<span>${linea}</span><span class="s">${total ? pct + '%' : ''}</span></div>
        ${total ? `<div class="track"><i style="width:${pct}%"></i></div>` : ''}
        ${ayuda ? `<p class="help">${ayuda}</p>` : ''}
      </div>` : ''}
      ${conError.length ? `<div class="note bad">${ic('alert')}<span>${conError.length} ${conError.length === 1 ? 'foto no se pudo subir' : 'fotos no se pudieron subir'}: ${esc(conError[0].error)} <button class="link" data-act="descartar">Descartarlas</button></span></div>` : ''}`;
  }
  function seccionDatos(d) {
    const hechos = datosHechos(d);
    return `<details class="sec" data-sec="datos" ${st.abiertos.datos ? 'open' : ''}>
      <summary><span>Datos del ${T(d).uno}</span><span class="prog ${hechos === CAMPOS.length ? 'ok' : ''}">${hechos} de ${CAMPOS.length}</span>${ic('chev')}</summary>
      <div class="sec-in">
        <div class="form-grid">${CAMPOS.map(([k, e, ej]) => campo(k, e, ej, d[k], 'det')).join('')}</div>
        <div class="kv"><span>Hora de inicio</span><b>${fmtDate(d.fecha)} · ${fmtTime(d.fecha)}</b><span>Hora de fin</span><b>${d.cerrado ? `${fmtDate(d.cerrado)} · ${fmtTime(d.cerrado)}` : 'Se guarda al cerrar'}</b></div>
        <p class="help">Los cambios se guardan solos. Transporte, carga y sellos se completan en el informe de Word.</p>
        ${sugerencias()}
      </div>
    </details>`;
  }
  function seccionInspeccion(d) {
    const hechos = inspHechos(d);
    const insp = d.inspeccion || {};
    return `<details class="sec" data-sec="insp" ${st.abiertos.insp ? 'open' : ''}>
      <summary><span>Inspección previa</span><span class="prog ${hechos === INSPECCION.length ? 'ok' : ''}">${hechos} de ${INSPECCION.length}</span>${ic('chev')}</summary>
      <div class="sec-in insp">
        ${INSPECCION.map(([k, nombre]) => {
          const r = insp[k] || {};
          return `<div class="insp-row">
            <span class="insp-n">${esc(nombre)}</span>
            <div class="seg" role="group" aria-label="${esc(nombre)}">
              <button data-act="insp" data-k="${k}" data-v="si" aria-pressed="${r.v === 'si'}">Sí</button>
              <button data-act="insp" data-k="${k}" data-v="no" aria-pressed="${r.v === 'no'}">No</button>
            </div>
            ${r.v === 'no' ? `<input class="input obs" data-obs="${k}" placeholder="¿Qué se encontró?" value="${esc(r.obs || '')}" maxlength="200" aria-label="Observación: ${esc(nombre)}">` : ''}
          </div>`;
        }).join('')}
        <p class="help">Se marca antes de cargar. Si algo está mal, escribe la observación; sale en el informe.</p>
      </div>
    </details>`;
  }
  function vDespacho() {
    const d = buscar(vista.id);
    if (!d) return vModulo();
    const t = T(d);
    const abierto = d.estado === 'proceso';
    const pend = pendientesDe(d.id);
    const borradores = pend.filter(x => !x.listo).length;
    const fotoNum = name => { const m = /(\d+)/.exec(name); return m ? m[1] : ''; };
    // Botón X para eliminar (solo con el registro abierto); al tocarlo pide confirmar sobre la misma foto.
    const quitar = (id, tipo, nombre) => {
      if (!abierto || id === subiendoId) return '';
      if (st.borrar === id) return `<div class="confirm"><span>¿Eliminar foto ${esc(fotoNum(nombre))}?</span>
        <button class="si" data-act="borrarfoto" data-id="${esc(id)}" data-tipo="${tipo}" ${st.ocupado ? 'disabled' : ''}>Eliminar</button>
        <button class="no" data-act="noborrar">Cancelar</button></div>`;
      return `<button class="del" data-act="preborrar" data-id="${esc(id)}" aria-label="Eliminar foto ${esc(fotoNum(nombre))}">${ic('x')}</button>`;
    };
    const tilesNube = vista.nube.slice().reverse().map(f => {
      const src = recientes.get(f.name) || f.thumb;
      return `<div class="tile">${src ? `<img src="${esc(src)}" alt="${esc(f.name)}" loading="lazy">` : `<div class="ph">${ic('image')}</div>`}<span class="num">${esc(fotoNum(f.name))}</span><span class="st ok" title="Subida">${ic('check')}</span>${quitar(f.id, 'nube', f.name)}</div>`;
    });
    const tilesCola = pend.slice().reverse().map(x => {
      const badge = x.error ? `<span class="st err" title="No se pudo subir">${ic('alert')}</span>`
        : !x.listo ? `<span class="st local" title="Sin subir">Sin subir</span>`
        : x.id === subiendoId ? `<span class="st up" title="Subiendo">${ic('spin')}</span>`
        : `<span class="st wait" title="Esperando">${ic('clock')}</span>`;
      return `<div class="tile ${x.listo ? 'pend' : 'draft'}"><img src="${urlsCola.get(x.id)}" alt="${esc(x.archivo)}"><span class="num">${esc(fotoNum(x.archivo))}</span>${badge}${quitar(x.id, 'cola', x.archivo)}</div>`;
    });
    const tiles = tilesCola.concat(tilesNube);
    return `<main class="wrap">
      <button class="btn ghost back" data-act="volver">${ic('back')}${st.prev === 'registros' ? 'Registros' : 'Inicio'}</button>
      <div class="d-head">
        <span class="eyebrow">${t.Uno}</span>
        <h1>${esc(d.nombre)}</h1>
        <div class="d-meta"><span class="pill ${d.estado}">${abierto ? 'En proceso' : 'Cerrado'}</span><span>${fmtDate(d.fecha)} · ${fmtTime(d.fecha)}${d.cliente ? ' · ' + esc(d.cliente) : ''}</span></div>
        <div class="crumbs">${ic('folder')}OneDrive › ${esc(RAIZ)} › ${esc(t.carpeta)} › <b>${esc(d.carpeta)}</b></div>
      </div>
      ${abierto ? `
        <div class="row">
          <button class="btn primary big" data-act="camara" style="flex:2 1 200px">${ic('camera')}Tomar foto</button>
          <button class="btn big" data-act="galeria" style="flex:1 1 140px">${ic('image')}Desde el celular</button>
        </div>` : `
        <div class="closed-note"><span>${t.Uno} cerrado el ${fmtDate(d.cerrado)} (${fmtTime(d.cerrado)}). Si faltó alguna foto, reábrelo para agregarla.</span>
          <button class="btn" data-act="reabrir" ${st.ocupado ? 'disabled' : ''}>${ic('redo')}Reabrir para agregar fotos</button></div>`}
      ${cajaSubida(d)}
      <section class="list">
        <div class="sec-h"><h2>Fotos</h2><span class="count">${tiles.length}</span></div>
        ${vista.error ? `<div class="note bad">${ic('alert')}<span>${esc(vista.error)} <button class="link" data-act="recargarvista">Reintentar</button></span></div>` : ''}
        ${tiles.length ? `<div class="grid">${tiles.join('')}</div>`
          : vista.cargando ? `<p class="loading">${ic('spin')}Cargando fotos…</p>`
          : `<div class="empty">Toca <b>Tomar foto</b> para empezar. Las fotos quedan en el celular hasta que tocas <b>Subir</b>.</div>`}
      </section>
      ${seccionDatos(d)}
      ${seccionInspeccion(d)}
      <div class="row">
        <button class="btn primary" data-act="informe" data-id="${esc(d.id)}" ${st.generando ? 'disabled' : ''}>${st.generando === d.id ? ic('spin') + 'Preparando informe…' : ic('doc') + 'Descargar informe (Word)'}</button>
        <a class="btn" href="${esc(d.webUrl)}" target="_blank" rel="noopener">${ic('ext')}Abrir carpeta en OneDrive</a>
      </div>
      ${abierto ? `<button class="btn block" data-act="cerrar" ${tiles.length && !borradores && !st.ocupado ? '' : 'disabled'}>${ic('check')}Cerrar ${t.uno}</button>
        ${borradores ? `<p class="help">Para cerrar, primero sube las ${borradores} ${borradores === 1 ? 'foto' : 'fotos'} que están en el celular o elimínalas.</p>` : ''}` : ''}
    </main>`;
  }

  function vRevision() {
    const r = st.revision;
    if (r.origen === 'camara') {
      const it = r.items[0];
      return `<div class="review" role="dialog" aria-label="Revisar foto">
        <div class="review-top"><b>¿La foto salió bien?</b><span class="mono">${fmtSize(it.orig)} → ${fmtSize(it.blob.size)}</span></div>
        <div class="review-img"><img src="${it.url}" alt="Foto tomada"></div>
        <p class="hint">Revisa que esté enfocada y se vea lo importante.</p>
        <div class="review-bar">
          <button class="btn sec" data-act="repetir">${ic('redo')}Volver a tomar</button>
          <button class="btn acc" data-act="aceptar">${ic('check')}Aceptar</button>
        </div></div>`;
    }
    const total = r.items.reduce((a, b) => a + b.blob.size, 0);
    return `<div class="review" role="dialog" aria-label="Revisar fotos">
      <div class="review-top"><b>${r.items.length} ${r.items.length === 1 ? 'foto seleccionada' : 'fotos seleccionadas'}</b><span class="mono">${fmtSize(total)}</span></div>
      <div class="review-grid"><div class="grid">${r.items.map((it, i) => `<div class="tile"><img src="${it.url}" alt="Foto ${i + 1}"><button class="rm" data-act="quitar" data-i="${i}" aria-label="Quitar foto">${ic('x')}</button></div>`).join('')}</div></div>
      <p class="hint">Quita las que no sirvan antes de agregarlas.</p>
      <div class="review-bar">
        <button class="btn sec" data-act="cancelar">Cancelar</button>
        <button class="btn acc" data-act="aceptar" ${r.items.length ? '' : 'disabled'}>${ic('check')}Agregar ${r.items.length}</button>
      </div></div>`;
  }

  function vMenu() {
    if (!st.menu) return '';
    const pend = cola.filter(x => !x.error).length;
    return `<div class="sheet-bg" data-act="cerrarmenu"><div class="sheet" data-stop>
      <span class="eyebrow">Sesión</span>
      <p><b>${esc(cuenta.name || '')}</b><br><span class="muted">${esc(cuenta.username)}</span></p>
      <p class="help">La sesión queda abierta en este dispositivo aunque cierres la app.</p>
      ${pend ? `<div class="note">${ic('alert')}<span>Hay ${pend} ${pend === 1 ? 'foto' : 'fotos'} sin subir. Súbelas y espera a que terminen antes de cerrar sesión.</span></div>` : ''}
      <button class="btn block danger" data-act="salir" ${pend ? 'disabled' : ''}>${ic('logout')}Cerrar sesión</button>
      <button class="btn block" data-act="cerrarmenu">Cancelar</button>
    </div></div>`;
  }

  function render() {
    const v = st.view;
    if (v === 'boot') return;
    if (v === 'ingreso') { app.innerHTML = vIngreso(); return; }
    if (v === 'sinacceso') { app.innerHTML = vSinAcceso(); return; }
    if (v === 'falla') { app.innerHTML = vFalla(); return; }
    const cuerpo = v === 'nuevo' ? vNuevo() : v === 'despacho' ? vDespacho() : v === 'registros' ? vRegistros() : v === 'modulo' ? vModulo() : vAreas();
    app.innerHTML = encabezado(v === 'registros' || v === 'areas') + cuerpo + barra() + (st.revision ? vRevision() : '') + vMenu();
  }

  // ---------- Vista de un registro ----------
  async function cargarVista(id) {
    vista = { id, nube: [], cargando: true, error: null };
    st.borrar = null;
    const d = buscar(id);
    if (d) st.tipo = d.tipo;
    render();
    try {
      const r = await graph(`/me/drive/items/${id}/children?$select=id,name,file&$expand=thumbnails&$top=999`);
      if (vista.id !== id) return;
      vista.nube = r.value.filter(f => f.file && ES_FOTO.test(f.name))
        .map(f => ({ id: f.id, name: f.name, thumb: f.thumbnails && f.thumbnails[0] && f.thumbnails[0].medium && f.thumbnails[0].medium.url }))
        .sort((a, b) => a.name.localeCompare(b.name, 'es', { numeric: true }));
      if (d && d.fotos !== vista.nube.length) cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === id); if (x) x.fotos = vista.nube.length; }).catch(() => {});
    } catch (e) {
      if (vista.id !== id) return;
      vista.error = e instanceof SinSenal ? 'Sin señal: no se pueden mostrar las fotos que ya están en OneDrive.' : 'No se pudieron cargar las fotos.';
    }
    vista.cargando = false;
    render();
  }

  // ---------- Eventos ----------
  cam.addEventListener('change', () => { tomar(cam.files, 'camara'); cam.value = ''; });
  gal.addEventListener('change', () => { tomar(gal.files, 'galeria'); gal.value = ''; });
  async function tomar(files, origen) {
    if (!files || !files.length) return;
    const items = await Promise.all([...files].map(comprimir));
    st.revision = { origen, items };
    render();
  }
  function soltarRevision() {
    if (st.revision) st.revision.items.forEach(it => URL.revokeObjectURL(it.url));
    st.revision = null;
  }
  const actualizarLista = () => cargarRegistro().then(() => {
    if (st.view === 'registros') document.getElementById('lista').innerHTML = listaRegistro();
    else refrescar();
  }).catch(() => {});
  function entrarArea(tipo) {
    if (tipo !== st.tipo) Object.assign(st, { q: '', desde: '', hasta: '', estado: 'todos' });
    st.tipo = tipo;
    ir('modulo');
    actualizarLista();
  }

  document.addEventListener('toggle', e => {
    const s = e.target.closest && e.target.closest('details[data-sec]');
    if (s) st.abiertos[s.dataset.sec] = s.open;
  }, true);

  document.addEventListener('click', async e => {
    if (DEMO && e.target.closest('a[target="_blank"]')) {
      e.preventDefault();
      toast('En la app real, este botón abre la carpeta en OneDrive.');
      return;
    }
    if (e.target.closest('[data-stop]') && !e.target.closest('[data-act]')) return;
    const b = e.target.closest('[data-act]');
    if (!b) return;
    const d = buscar(vista.id);
    switch (b.dataset.act) {
      case 'ingresar':
        if (DEMO) { st.opcionesDemo = true; render(); } else ingresar();
        break;
      case 'demologin': iniciarDemo(b.dataset.u); break;
      case 'salir':
        if (DEMO) { try { localStorage.removeItem(SESION_DEMO); } catch (_) {} location.reload(); } else salir();
        break;
      case 'recargar': location.reload(); break;
      case 'menu': st.menu = true; render(); break;
      case 'cerrarmenu': if (e.target === b || b.tagName === 'BUTTON') { st.menu = false; render(); } break;
      case 'areas': ir('areas'); actualizarLista(); break;
      case 'area': entrarArea(b.dataset.t); break;
      case 'ir': ir(b.dataset.v); actualizarLista(); break;
      case 'registros': ir('registros'); actualizarLista(); break;
      case 'volver': ir(st.prev === 'registros' ? 'registros' : 'modulo'); actualizarLista(); break;
      case 'nuevo': {
        // El lugar de carga casi siempre se repite: se propone el del último registro.
        const ultimo = ordenados(st.tipo)[0];
        st.nuevo = { lugarCarga: ultimo ? ultimo.lugarCarga || '' : '' };
        ir('nuevo');
        setTimeout(() => document.getElementById('nuevo-nombre')?.focus(), 30);
        break;
      }
      case 'abrir': {
        st.prev = st.view === 'registros' ? 'registros' : 'modulo';
        st.view = 'despacho'; st.menu = false;
        cargarVista(b.dataset.id);
        window.scrollTo(0, 0);
        break;
      }
      case 'crear': {
        if (st.ocupado || (st.nuevo.nombre || '').trim().length < 3) return;
        st.ocupado = true; render();
        try {
          const nd = await crearRegistro(st.tipo, st.nuevo);
          st.ocupado = false;
          st.prev = 'modulo';
          st.abiertos = {};
          vista = { id: nd.id, nube: [], cargando: false, error: null };
          ir('despacho');
          toast(DEMO ? 'Carpeta creada' : 'Carpeta creada en OneDrive');
        } catch (err) {
          st.ocupado = false; render();
          toast(errorAMensaje(err, 'crear la carpeta'));
        }
        break;
      }
      case 'informe': {
        const x = buscar(b.dataset.id);
        if (x && !st.generando) descargarInforme(x);
        break;
      }
      case 'insp': {
        const k = b.dataset.k, v = b.dataset.v;
        const actual = d.inspeccion && d.inspeccion[k] && d.inspeccion[k].v;
        guardarDetalles(d.id, { inspeccion: { [k]: { v: actual === v ? null : v } } });
        render();
        if (v === 'no' && actual !== 'no') setTimeout(() => document.querySelector(`[data-obs="${k}"]`)?.focus(), 30);
        break;
      }
      case 'camara': cam.click(); break;
      case 'galeria': gal.click(); break;
      case 'repetir': soltarRevision(); render(); cam.click(); break;
      case 'cancelar': soltarRevision(); render(); break;
      case 'quitar': {
        const [it] = st.revision.items.splice(+b.dataset.i, 1);
        URL.revokeObjectURL(it.url); render();
        break;
      }
      case 'aceptar': {
        const items = st.revision.items;
        st.revision = null;
        items.forEach(it => URL.revokeObjectURL(it.url));
        const n = await encolar(items, d);
        render();
        toast(items.length === 1 ? `Foto ${pad3(n)} agregada. Súbela cuando termines.` : `${items.length} fotos agregadas. Súbelas cuando termines.`);
        break;
      }
      case 'subirfotos': {
        const n = await subirFotos(d);
        render();
        toast(sinSenal ? `${n} ${n === 1 ? 'foto' : 'fotos'} en espera: se subirán cuando haya señal.` : `Subiendo ${n} ${n === 1 ? 'foto' : 'fotos'}…`);
        break;
      }
      case 'descartar': {
        for (const x of pendientesDe(d.id).filter(x => x.error)) { await idb.quitar(x.id); urlsCola.delete(x.id); }
        await refrescarCola(); render();
        break;
      }
      case 'recargarvista': cargarVista(vista.id); break;
      case 'preborrar': st.borrar = b.dataset.id; render(); break;
      case 'noborrar': st.borrar = null; render(); break;
      case 'borrarfoto': {
        const id = b.dataset.id;
        if (b.dataset.tipo === 'cola') {
          if (id === subiendoId) { toast('Esa foto se está subiendo; espera un momento.'); break; }
          await idb.quitar(id);
          urlsCola.delete(id);
          await refrescarCola();
          st.borrar = null; render();
          toast('Foto eliminada');
          break;
        }
        st.ocupado = true; render();
        try {
          await graph(`/me/drive/items/${id}`, { method: 'DELETE' });
          vista.nube = vista.nube.filter(f => f.id !== id);
          const n = vista.nube.length;
          d.fotos = n;
          cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); if (x) x.fotos = n; }).catch(() => {});
          toast(DEMO ? 'Foto eliminada' : 'Foto eliminada. Si fue un error, está en la papelera de OneDrive.');
        } catch (err) {
          toast(err.status === 404 ? 'Esa foto ya no estaba en la carpeta.' : errorAMensaje(err, 'eliminar la foto'));
          if (err.status === 404) vista.nube = vista.nube.filter(f => f.id !== id);
        }
        st.ocupado = false; st.borrar = null; render();
        break;
      }
      case 'cerrar': {
        if (porSubir(d.id).length) { toast('Primero sube las fotos que están en el celular.'); break; }
        st.ocupado = true; render();
        try {
          const n = await contarFotos(d.id);
          await cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); x.estado = 'cerrado'; x.cerrado = new Date().toISOString(); x.fotos = Math.max(n, x.fotos || 0); });
          st.ocupado = false;
          ir('modulo');
          toast(`${T(d).Uno} "${d.nombre}" cerrado`);
        } catch (err) {
          st.ocupado = false; render();
          toast(errorAMensaje(err, `cerrar el ${T(d).uno}`));
        }
        break;
      }
      case 'reabrir': {
        st.ocupado = true; render();
        try {
          await cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); x.estado = 'proceso'; x.cerrado = null; });
          toast(`${T(d).Uno} reabierto`);
        } catch (err) { toast(errorAMensaje(err, `reabrir el ${T(d).uno}`)); }
        st.ocupado = false; render();
        break;
      }
      case 'estado': st.estado = b.dataset.e; render(); break;
      case 'limpiar': Object.assign(st, { q: '', desde: '', hasta: '', estado: 'todos' }); render(); break;
    }
  });

  document.addEventListener('input', e => {
    const t = e.target;
    if (t.id && t.id.startsWith('nuevo-')) {
      st.nuevo[t.dataset.campo] = t.value;
      if (t.dataset.campo === 'nombre') {
        const { nombre, quitados } = nombreCarpeta(t.value, new Date());
        document.getElementById('ruta').innerHTML = vistaRuta(nombre, quitados);
      }
      document.getElementById('crear').disabled = st.ocupado || (st.nuevo.nombre || '').trim().length < 3;
    } else if (t.id === 'q' || t.id === 'desde' || t.id === 'hasta') {
      st[t.id] = t.value;
      document.getElementById('lista').innerHTML = listaRegistro();
    }
  });
  // Datos y observaciones del registro abierto: se guardan al salir del campo.
  document.addEventListener('change', e => {
    const t = e.target;
    const d = buscar(vista.id);
    if (!d || st.view !== 'despacho') return;
    if (t.id && t.id.startsWith('det-')) {
      const v = t.value.trim();
      if ((d[t.dataset.campo] || '') !== v) { guardarDetalles(d.id, { [t.dataset.campo]: v }); toast('Guardado'); }
    } else if (t.dataset.obs) {
      guardarDetalles(d.id, { inspeccion: { [t.dataset.obs]: { obs: t.value.trim() } } });
      toast('Guardado');
    }
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.target.id && e.target.id.startsWith('nuevo-')) document.getElementById('crear').click();
    if (e.key === 'Enter' && (e.target.id?.startsWith('det-') || e.target.dataset?.obs)) e.target.blur();
    if (e.key === 'Escape' && st.menu) { st.menu = false; render(); }
  });

  // ---------- Arranque ----------
  async function iniciarDemo(clave) {
    cuenta = DEMO_CUENTAS[clave];
    if (!cuenta) { ir('ingreso'); return; }
    if (clave === 'otro') { ir('sinacceso'); return; }
    try { localStorage.setItem(SESION_DEMO, clave); } catch (_) {}
    st.rol = clave;
    await demoCargar();
    await cargarRegistro();
    await refrescarCola();
    ir('areas');
    procesarCola();
  }

  async function arrancar() {
    if (DEMO) {
      let guardada = null;
      try { guardada = localStorage.getItem(SESION_DEMO); } catch (_) {}
      if (guardada && DEMO_CUENTAS[guardada]) await iniciarDemo(guardada);
      else ir('ingreso');
      return;
    }
    if (!window.msal) { ir('falla', { error: 'No se pudo cargar el inicio de sesión de Microsoft. Revisa la conexión a internet.' }); return; }
    if (!CFG.clientId || /PEGAR/.test(CFG.clientId + CFG.tenantId)) { ir('falla', { error: 'Falta configurar la app (config.js). Avisa a soporte.' }); return; }
    try { await iniciarSesionMsal(); }
    catch (e) { console.error(e); ir('falla', { error: 'No se pudo iniciar sesión con Microsoft. Intenta de nuevo.' }); return; }
    if (!cuenta) { ir('ingreso'); return; }

    const permitidos = {};
    Object.keys(CFG.correosPermitidos || {}).forEach(k => { permitidos[k.toLowerCase()] = CFG.correosPermitidos[k]; });
    st.rol = permitidos[(cuenta.username || '').toLowerCase()] || null;
    if (!st.rol) { ir('sinacceso'); return; }

    registro = aplicarPendientes(cache.leer('registro.' + cuenta.username) || { despachos: [] });
    await refrescarCola();
    ir('areas');
    try { await cargarRegistro(); refrescar(); }
    catch (e) { if (!(e instanceof SinSenal)) { console.error(e); toast('No se pudo actualizar el registro desde OneDrive.'); } }
    procesarCola();
  }

  if (!window.HUB_DEMO && 'serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  arrancar();
})();
