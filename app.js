/* Hub Comex — app de Ventas para fotografiar despachos y guardarlos en OneDrive.
 *
 * - Ingreso: Microsoft (MSAL). La sesión queda guardada en el dispositivo.
 * - Fotos: OneDrive de la cuenta que ingresa → /<carpetaRaiz>/<AAAA-MM-DD - nombre>/Foto_001.jpg
 * - Cola de subida: IndexedDB. Si no hay señal, las fotos esperan en el celular.
 * - Registro de despachos: archivo JSON dentro de la carpeta raíz (/<carpetaRaiz>/_registro-hub-comex.json).
 */
(function () {
  'use strict';

  const CFG = window.HUB_CONFIG || {};
  const SCOPES = CFG.permisos || ['User.Read', 'Files.ReadWrite'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const REGISTRO = `/me/drive/root:/${encodeURIComponent(CFG.carpetaRaiz || 'Despachos')}/_registro-hub-comex.json`;
  const LIMITE_SIMPLE = 4 * 1024 * 1024; // Graph acepta PUT directo hasta 4 MB
  // Modo prueba (config modo: 'prueba', o …/#demo): sin Microsoft. Se elige Ventas o Soporte al ingresar
  // y los despachos y fotos se guardan solo en este dispositivo.
  const DEMO = CFG.modo !== 'microsoft' || location.hash === '#demo' || window.HUB_DEMO === true;
  const DEMO_CUENTAS = {
    ventas: { username: 'ventas-sales@healthyfood.com.ec', name: 'Ventas' },
    soporte: { username: 'notificaciones@healthyfood.com.ec', name: 'Soporte' },
    otro: { username: 'otro.usuario@healthyfood.com.ec', name: 'Otro usuario' }
  };
  const SESION_DEMO = 'hc.demo.sesion';

  // ---------- Utilidades ----------
  const ICONS = {
    plus: '<path d="M12 5v14M5 12h14"/>',
    camera: '<path d="M4 8h3l2-3h6l2 3h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="13.5" r="3.6"/>',
    image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="1.8"/><path d="m21 16-5-5-9 9"/>',
    folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    ext: '<path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.2-4.2"/>',
    home: '<path d="M4 11 12 4l8 7v8a1 1 0 0 1-1 1h-5v-6h-4v6H5a1 1 0 0 1-1-1z"/>',
    list: '<path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01"/>',
    back: '<path d="M15 18l-6-6 6-6"/>',
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

  // ---------- Almacenamiento local (cola de fotos) ----------
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
    guardar(k, v) { try { localStorage.setItem('hc.' + k, JSON.stringify(v)); } catch (_) {} }
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

  // ---------- OneDrive simulado (solo modo demostración) ----------
  const demo = { carpetas: new Map(), raiz: null, registro: null, n: 0 };
  function demoError(status) { const e = new Error('Demo ' + status); e.status = status; return e; }
  function demoCarpeta(nombre, padre) {
    const id = 'demo-' + uid() + (++demo.n);
    const hermanos = padre ? [...demo.carpetas.values()].filter(c => c.padre === padre).map(c => c.name) : [];
    let name = nombre, k = 1;
    while (hermanos.includes(name)) name = `${nombre} ${k++}`;
    demo.carpetas.set(id, { id, name, padre, fotos: [] });
    return { id, name, webUrl: '#' };
  }
  function demoSemilla() {
    demo.raiz = demoCarpeta(CFG.carpetaRaiz || 'Despachos', null).id;
    const ejemplos = [['Pulpa mango Goya contenedor C6', 3, 38], ['Puré de banano orgánico contenedor B12', 6, 41],
      ['Pulpa maracuyá contenedor M3', 11, 29], ['Pulpa guanábana Goya contenedor G4', 18, 33]];
    const despachos = ejemplos.map(([nombre, dias, fotos]) => {
      const fecha = new Date(Date.now() - dias * 864e5);
      const c = demoCarpeta(nombreCarpeta(nombre, fecha).nombre, demo.raiz);
      demo.carpetas.get(c.id).fotos = Array.from({ length: fotos }, (_, i) => ({ id: c.id + '-' + i, name: `Foto_${pad3(i + 1)}.jpg` }));
      return { id: c.id, nombre, carpeta: c.name, webUrl: '#', fecha: fecha.toISOString(), estado: 'cerrado', cerrado: fecha.toISOString(), fotos };
    });
    demo.registro = JSON.stringify({ despachos });
  }
  // Lo simulado se guarda en este dispositivo: datos en localStorage, fotos en IndexedDB.
  const demoClave = () => 'demo.datos.' + cuenta.username;
  function demoGuardar() {
    cache.guardar(demoClave(), {
      n: demo.n, raiz: demo.raiz, registro: demo.registro,
      carpetas: [...demo.carpetas.values()].map(c => ({ id: c.id, name: c.name, padre: c.padre, fotos: c.fotos.map(f => ({ id: f.id, name: f.name })) }))
    });
  }
  async function demoCargar() {
    const d = cache.leer(demoClave());
    if (!d) { demoSemilla(); demoGuardar(); return; }
    Object.assign(demo, { n: d.n, raiz: d.raiz, registro: d.registro, carpetas: new Map() });
    for (const c of d.carpetas) {
      for (const f of c.fotos) {
        const blob = await idb.leer('demofoto.' + f.id).catch(() => null);
        if (blob) f.url = URL.createObjectURL(blob);
      }
      demo.carpetas.set(c.id, c);
    }
  }
  async function demoGraph(ruta, op) {
    await new Promise(r => setTimeout(r, 250 + Math.random() * 500));
    if (!navigator.onLine) throw new SinSenal();
    const metodo = op.method || 'GET';
    let m;
    if (ruta.startsWith(REGISTRO)) {
      if (metodo === 'PUT') { demo.registro = op.body; demoGuardar(); return {}; }
      if (!demo.registro) throw demoError(404);
      return { '@microsoft.graph.downloadUrl': 'data:application/json,' + encodeURIComponent(demo.registro) };
    }
    if (ruta.startsWith('/me/drive/root:/')) return { id: demo.raiz };
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

  // ---------- Registro de despachos ----------
  let registro = { despachos: [] };
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
    return data;
  }
  async function cargarRegistro() {
    registro = await leerRegistro();
    cache.guardar('registro.' + cuenta.username, registro);
    return registro;
  }
  // Lee lo último de OneDrive, aplica el cambio y lo guarda (evita pisar cambios hechos desde otro equipo).
  async function cambiarRegistro(cambio) {
    const data = await leerRegistro();
    cambio(data);
    await graph(REGISTRO + ':/content', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data, null, 1) });
    registro = data;
    cache.guardar('registro.' + cuenta.username, registro);
  }
  const buscar = id => registro.despachos.find(d => d.id === id);

  let idRaiz = null;
  async function carpetaRaiz() {
    if (idRaiz) return idRaiz;
    try {
      idRaiz = (await graph(`/me/drive/root:/${encodeURIComponent(CFG.carpetaRaiz)}`)).id;
    } catch (e) {
      if (e.status !== 404) throw e;
      idRaiz = (await graph('/me/drive/root/children', { method: 'POST', json: { name: CFG.carpetaRaiz, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' } })).id;
    }
    return idRaiz;
  }

  async function crearDespacho(texto) {
    const fecha = new Date();
    const { nombre } = nombreCarpeta(texto, fecha);
    const raiz = await carpetaRaiz();
    // "rename": si ya existe una carpeta con ese nombre, OneDrive agrega un número al final
    const f = await graph(`/me/drive/items/${raiz}/children`, { method: 'POST', json: { name: nombre, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' } });
    const d = {
      id: f.id, nombre: texto.trim().replace(/\s+/g, ' '), carpeta: f.name, webUrl: f.webUrl,
      fecha: fecha.toISOString(), estado: 'proceso', cerrado: null, fotos: 0
    };
    await cambiarRegistro(r => r.despachos.unshift(d));
    return d;
  }

  async function contarFotos(id) {
    const f = await graph(`/me/drive/items/${id}?$select=id,folder`);
    return (f.folder && f.folder.childCount) || 0;
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

  async function procesarCola() {
    if (trabajando || !cuenta) return;
    trabajando = true;
    clearTimeout(reintento);
    try {
      await refrescarCola();
      for (;;) {
        const it = cola.find(x => !x.error);
        if (!it) break;
        subiendoId = it.id; refrescar();
        try {
          const subido = await subirArchivo(it);
          recientes.set(subido.name, urlsCola.get(it.id));
          urlsCola.delete(it.id);
          await idb.quitar(it.id);
          sinSenal = false;
          if (vista.id === it.despachoId) vista.nube.push({ id: subido.id, name: subido.name });
        } catch (e) {
          if (e instanceof SinSenal) { sinSenal = true; reintento = setTimeout(procesarCola, 20000); break; }
          if (e.status === 404) { it.error = 'La carpeta del despacho ya no existe en OneDrive.'; await idb.poner(it); }
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
        archivo: `Foto_${pad3(n)}.${it.ext}`, blob: it.blob, creado: Date.now()
      });
    }
    await idb.guardar('contador.' + d.id, n);
    await refrescarCola();
    procesarCola();
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

  // ---------- Estado de la pantalla ----------
  const app = document.getElementById('app');
  const cam = document.getElementById('cam');
  const gal = document.getElementById('gal');
  const st = {
    view: 'boot', prev: 'home', rol: null, menu: false, revision: null, ocupado: false,
    nombre: '', q: '', desde: '', hasta: '', estado: 'todos', error: null
  };
  let vista = { id: null, nube: [], cargando: false, error: null };

  function toast(msg) {
    document.querySelector('.toast')?.remove();
    const t = document.createElement('div');
    t.className = 'toast'; t.setAttribute('role', 'status'); t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 3500);
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
  // Re-pinta solo pantallas sin campos de texto, para no interrumpir lo que se está escribiendo.
  function refrescar() {
    if (!st.revision && (st.view === 'despacho' || st.view === 'home')) render();
  }
  const saludo = () => { const h = new Date().getHours(); return h < 12 ? 'Buenos días' : h < 19 ? 'Buenas tardes' : 'Buenas noches'; };
  const ordenados = () => registro.despachos.slice().sort((a, b) => b.fecha.localeCompare(a.fecha));

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
    return `<header class="top ${ancho ? 'wide' : ''}"><div class="top-in">
      <div class="mark">HC</div><div class="top-name">Despachos<small>Área de Ventas</small></div>
      <button class="who" data-act="menu" aria-label="Cuenta"><span class="ava">${esc(nombre.trim()[0] || 'V').toUpperCase()}</span><span>${esc(nombre.split(' ')[0])}</span></button>
    </div></header>`;
  }
  function barra() {
    const v = st.view;
    return `<nav class="nav" aria-label="Principal"><div class="nav-in">
      <button class="nav-b" data-act="ir" data-v="home" ${v === 'home' ? 'aria-current="page"' : ''}>${ic('home')}Inicio</button>
      <button class="nav-plus" data-act="ir" data-v="nuevo" aria-label="Nuevo despacho">${ic('plus')}</button>
      <button class="nav-b" data-act="ir" data-v="registro" ${v === 'registro' ? 'aria-current="page"' : ''}>${ic('list')}Registro</button>
    </div></nav>`;
  }
  function tarjeta(d) {
    const pend = pendientesDe(d.id).length;
    return `<div class="card">
      <button class="main" data-act="abrir" data-id="${esc(d.id)}">
        <span class="name">${esc(d.nombre)}</span>
        <span class="meta">${fmtDate(d.fecha)} · ${d.fotos} fotos${pend ? ` · ${pend} por subir` : ''} <span class="pill ${d.estado}">${d.estado === 'proceso' ? 'En proceso' : 'Cerrado'}</span></span>
      </button>
      <a class="icon-btn" href="${esc(d.webUrl)}" target="_blank" rel="noopener" aria-label="Abrir carpeta en OneDrive" title="Abrir carpeta">${ic('folder')}</a>
    </div>`;
  }

  function vInicio() {
    const lista = ordenados();
    const enProceso = lista.filter(d => d.estado === 'proceso');
    const hoy = new Date();
    const delMes = lista.filter(d => { const f = new Date(d.fecha); return f.getMonth() === hoy.getMonth() && f.getFullYear() === hoy.getFullYear(); });
    const pend = cola.filter(x => !x.error).length;
    return `<main class="wrap">
      <div class="greet"><h1>${saludo()},<br>${st.rol === 'soporte' ? 'Soporte' : 'equipo de Ventas'}</h1>
        <p>${hoy.toLocaleDateString('es-EC', { weekday: 'long', day: 'numeric', month: 'long' })} · ${delMes.length} ${delMes.length === 1 ? 'despacho' : 'despachos'} este mes</p></div>
      ${DEMO ? `<div class="note info">${ic('info')}<span><b>Versión de prueba.</b> Los despachos y fotos se guardan solo en este dispositivo; todavía no se suben a OneDrive.</span></div>` : ''}
      ${st.rol === 'soporte' ? `<div class="note info">${ic('info')}<span>Estás en modo soporte. Tus pruebas se guardan en tu propio OneDrive, no en el de Ventas.</span></div>` : ''}
      ${pend ? `<div class="note ${sinSenal ? 'bad' : ''}">${ic(sinSenal ? 'nosignal' : 'cloud')}<span>${pend} ${pend === 1 ? 'foto' : 'fotos'} ${sinSenal ? 'esperando señal para subir.' : 'subiendo a OneDrive…'}</span></div>` : ''}
      <button class="cta" data-act="ir" data-v="nuevo"><span class="plus">${ic('plus')}</span><span><b>Nuevo despacho</b><small>Ponle nombre y empieza a tomar fotos</small></span></button>
      ${enProceso.length ? `<div class="list"><span class="eyebrow">Sin cerrar</span>${enProceso.map(tarjeta).join('')}</div>` : ''}
      <section class="list">
        <div class="sec-h"><h2>Últimos despachos</h2>${lista.length ? `<button class="link" data-act="ir" data-v="registro">Ver registro</button>` : ''}</div>
        ${lista.length ? lista.filter(d => d.estado !== 'proceso').slice(0, 3).map(tarjeta).join('') || '<p class="help">Los despachos cerrados aparecerán aquí.</p>'
          : `<div class="empty">Aún no hay despachos. Toca <b>Nuevo despacho</b> para crear el primero.</div>`}
      </section>
    </main>`;
  }

  function vistaRuta(final, quitados) {
    return `<span class="eyebrow">Se creará la carpeta</span>
      <div class="crumbs">${ic('folder')}OneDrive <span>›</span> ${esc(CFG.carpetaRaiz)} <span>›</span> <b>${final ? esc(final) : '…'}</b></div>
      ${quitados.length ? `<p class="warn">${ic('alert')}<span>OneDrive no acepta ${quitados.map(c => `<code class="mono">${esc(c)}</code>`).join(' ')} en nombres de carpeta; se cambiaron por guiones.</span></p>` : ''}`;
  }
  function vNuevo() {
    const { nombre, quitados } = nombreCarpeta(st.nombre, new Date());
    return `<main class="wrap">
      <button class="btn ghost back" data-act="ir" data-v="home">${ic('back')}Inicio</button>
      <h1 style="font-size:28px">Nuevo despacho</h1>
      <div class="field">
        <label for="nombre">Nombre del despacho</label>
        <input class="input" id="nombre" autocomplete="off" placeholder="Ej. Pulpa mango Goya contenedor C7" value="${esc(st.nombre)}" maxlength="120">
        <p class="help">Será el nombre de la carpeta. Escríbelo como lo vas a buscar después.</p>
      </div>
      <div class="preview-path" id="ruta">${vistaRuta(nombre, quitados)}</div>
      <button class="btn primary big block" data-act="crear" id="crear" ${st.nombre.trim().length >= 3 && !st.ocupado ? '' : 'disabled'}>${st.ocupado ? ic('spin') + 'Creando carpeta…' : ic('camera') + 'Crear y tomar fotos'}</button>
    </main>`;
  }

  function cajaSubida(d) {
    const pend = pendientesDe(d.id);
    const conError = pend.filter(x => x.error);
    const esperando = pend.length - conError.length;
    const subidas = vista.nube.length;
    const total = subidas + esperando;
    const pct = total ? Math.round(subidas / total * 100) : 0;
    let cls = '', icon = 'cloud', linea, ayuda;
    if (vista.cargando && !total) { linea = 'Revisando fotos en OneDrive…'; ayuda = ''; }
    else if (!total) { linea = 'Aún no hay fotos'; ayuda = 'Cada foto que aceptes se sube sola a la carpeta.'; }
    else if (!esperando) { cls = 'ok'; icon = 'cloudok'; linea = `${total} de ${total} fotos subidas`; ayuda = DEMO ? 'Todas las fotos están guardadas.' : 'Todas las fotos ya están en OneDrive.'; }
    else if (sinSenal) { cls = 'off'; icon = 'nosignal'; linea = `${subidas} de ${total} fotos subidas`; ayuda = `<b>Sin señal.</b> ${esperando} ${esperando === 1 ? 'foto espera' : 'fotos esperan'} en el celular y se ${esperando === 1 ? 'subirá' : 'subirán'} solas cuando vuelva la conexión. Puedes seguir tomando fotos.`; }
    else { cls = 'going'; icon = 'cloud'; linea = `${subidas} de ${total} fotos subidas`; ayuda = 'Subiendo en segundo plano. Puedes seguir tomando fotos; no cierres la app hasta que terminen.'; }
    return `<div class="upbox ${cls}">
      <div class="upline">${ic(icon)}<span>${linea}</span><span class="s">${total ? pct + '%' : ''}</span></div>
      ${total ? `<div class="track"><i style="width:${pct}%"></i></div>` : ''}
      ${ayuda ? `<p class="help">${ayuda}</p>` : ''}
      ${conError.length ? `<div class="note bad">${ic('alert')}<span>${conError.length} ${conError.length === 1 ? 'foto no se pudo subir' : 'fotos no se pudieron subir'}: ${esc(conError[0].error)} <button class="link" data-act="descartar">Descartarlas</button></span></div>` : ''}
    </div>`;
  }
  function vDespacho() {
    const d = buscar(vista.id);
    if (!d) return vInicio();
    const abierto = d.estado === 'proceso';
    const pend = pendientesDe(d.id);
    const fotoNum = name => { const m = /(\d+)/.exec(name); return m ? m[1] : ''; };
    // Botón X para eliminar (solo con el despacho abierto); al tocarlo pide confirmar sobre la misma foto.
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
        : x.id === subiendoId ? `<span class="st up" title="Subiendo">${ic('spin')}</span>`
        : `<span class="st wait" title="Esperando">${ic('clock')}</span>`;
      return `<div class="tile pend"><img src="${urlsCola.get(x.id)}" alt="${esc(x.archivo)}"><span class="num">${esc(fotoNum(x.archivo))}</span>${badge}${quitar(x.id, 'cola', x.archivo)}</div>`;
    });
    const tiles = tilesCola.concat(tilesNube);
    return `<main class="wrap">
      <button class="btn ghost back" data-act="ir" data-v="${st.prev}">${ic('back')}${st.prev === 'registro' ? 'Registro' : 'Inicio'}</button>
      <div class="d-head">
        <h1>${esc(d.nombre)}</h1>
        <div class="d-meta"><span class="pill ${d.estado}">${abierto ? 'En proceso' : 'Cerrado'}</span><span>${fmtDate(d.fecha)} · ${fmtTime(d.fecha)}</span></div>
        <div class="crumbs">${ic('folder')}OneDrive › ${esc(CFG.carpetaRaiz)} › <b>${esc(d.carpeta)}</b></div>
      </div>
      ${abierto ? `
        <div class="row">
          <button class="btn primary big" data-act="camara" style="flex:2 1 200px">${ic('camera')}Tomar foto</button>
          <button class="btn big" data-act="galeria" style="flex:1 1 140px">${ic('image')}Desde el celular</button>
        </div>` : `
        <div class="closed-note"><span>Despacho cerrado el ${fmtDate(d.cerrado)}. Si faltó alguna foto, reábrelo para agregarla.</span>
          <button class="btn" data-act="reabrir" ${st.ocupado ? 'disabled' : ''}>${ic('redo')}Reabrir para agregar fotos</button></div>`}
      ${cajaSubida(d)}
      <section class="list">
        <div class="sec-h"><h2>Fotos</h2><span class="count">${tiles.length}</span></div>
        ${vista.error ? `<div class="note bad">${ic('alert')}<span>${esc(vista.error)} <button class="link" data-act="recargarvista">Reintentar</button></span></div>` : ''}
        ${tiles.length ? `<div class="grid">${tiles.join('')}</div>`
          : vista.cargando ? `<p class="loading">${ic('spin')}Cargando fotos…</p>`
          : `<div class="empty">Toca <b>Tomar foto</b> para empezar. Cada foto la revisas antes de que se suba.</div>`}
      </section>
      <div class="row">
        <a class="btn" href="${esc(d.webUrl)}" target="_blank" rel="noopener">${ic('ext')}Abrir carpeta en OneDrive</a>
        ${abierto ? `<button class="btn" data-act="cerrar" ${tiles.length && !st.ocupado ? '' : 'disabled'}>${ic('check')}Cerrar despacho</button>` : ''}
      </div>
      ${abierto && pend.length ? `<p class="help">Si cierras el despacho ahora, las ${pend.length} fotos pendientes se suben igual cuando haya señal.</p>` : ''}
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
      <p class="hint">Quita las que no sirvan antes de subir.</p>
      <div class="review-bar">
        <button class="btn sec" data-act="cancelar">Cancelar</button>
        <button class="btn acc" data-act="aceptar" ${r.items.length ? '' : 'disabled'}>${ic('cloud')}Subir ${r.items.length}</button>
      </div></div>`;
  }

  function filtrados() {
    const q = st.q.trim().toLowerCase();
    return ordenados().filter(d => {
      if (q && !d.nombre.toLowerCase().includes(q) && !d.carpeta.toLowerCase().includes(q)) return false;
      const dia = isoDay(new Date(d.fecha));
      if (st.desde && dia < st.desde) return false;
      if (st.hasta && dia > st.hasta) return false;
      if (st.estado !== 'todos' && d.estado !== st.estado) return false;
      return true;
    });
  }
  function listaRegistro() {
    const filas = filtrados();
    const hayFiltro = st.q || st.desde || st.hasta || st.estado !== 'todos';
    return `<div class="sec-h"><span class="count">${filas.length} ${filas.length === 1 ? 'despacho' : 'despachos'}</span>
        ${hayFiltro ? `<button class="link" data-act="limpiar">Quitar filtros</button>` : ''}</div>
      ${filas.length ? filas.map(tarjeta).join('')
        : `<div class="empty">${registro.despachos.length ? 'Ningún despacho coincide con la búsqueda.' : 'Aún no hay despachos registrados.'}</div>`}`;
  }
  function vRegistro() {
    const chip = (v, l) => `<button class="chip" data-act="estado" data-e="${v}" aria-pressed="${st.estado === v}">${l}</button>`;
    return `<main class="wrap wide">
      <h1 style="font-size:28px">Registro de despachos</h1>
      <div class="reg-layout">
        <div class="filters">
          <div class="search">${ic('search')}<input class="input" id="q" type="search" placeholder="Buscar por nombre o contenedor" value="${esc(st.q)}" autocomplete="off"></div>
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

  function vMenu() {
    if (!st.menu) return '';
    const pend = cola.filter(x => !x.error).length;
    return `<div class="sheet-bg" data-act="cerrarmenu"><div class="sheet" data-stop>
      <span class="eyebrow">Sesión</span>
      <p><b>${esc(cuenta.name || '')}</b><br><span class="muted">${esc(cuenta.username)}</span></p>
      <p class="help">La sesión queda abierta en este dispositivo aunque cierres la app.</p>
      ${pend ? `<div class="note">${ic('alert')}<span>Hay ${pend} ${pend === 1 ? 'foto' : 'fotos'} sin subir. Espera a que terminen antes de cerrar sesión.</span></div>` : ''}
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
    const cuerpo = v === 'nuevo' ? vNuevo() : v === 'despacho' ? vDespacho() : v === 'registro' ? vRegistro() : vInicio();
    app.innerHTML = encabezado(v === 'registro') + cuerpo + barra() + (st.revision ? vRevision() : '') + vMenu();
  }

  // ---------- Vista de un despacho ----------
  async function cargarVista(id) {
    vista = { id, nube: [], cargando: true, error: null };
    st.borrar = null;
    render();
    try {
      const r = await graph(`/me/drive/items/${id}/children?$select=id,name,file&$expand=thumbnails&$top=999`);
      if (vista.id !== id) return;
      vista.nube = r.value.filter(f => f.file)
        .map(f => ({ id: f.id, name: f.name, thumb: f.thumbnails && f.thumbnails[0] && f.thumbnails[0].medium && f.thumbnails[0].medium.url }))
        .sort((a, b) => a.name.localeCompare(b.name, 'es', { numeric: true }));
      const d = buscar(id);
      if (d && d.fotos !== vista.nube.length) cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === id); if (x) x.fotos = vista.nube.length; }).catch(() => {});
    } catch (e) {
      if (vista.id !== id) return;
      vista.error = e instanceof SinSenal ? 'Sin señal: no se pueden mostrar las fotos que ya están en OneDrive.' : 'No se pudieron cargar las fotos de OneDrive.';
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

  document.addEventListener('click', async e => {
    if (DEMO && e.target.closest('a[target="_blank"]')) {
      e.preventDefault();
      toast('En la app real, este botón abre la carpeta del despacho en OneDrive.');
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
      case 'ir': {
        const v = b.dataset.v;
        if (v === 'nuevo') st.nombre = '';
        ir(v);
        if (v === 'nuevo') setTimeout(() => document.getElementById('nombre')?.focus(), 30);
        if (v === 'home' || v === 'registro') cargarRegistro().then(refrescar).catch(() => {});
        break;
      }
      case 'abrir': {
        st.prev = st.view === 'registro' ? 'registro' : 'home';
        st.view = 'despacho'; st.menu = false;
        cargarVista(b.dataset.id);
        window.scrollTo(0, 0);
        break;
      }
      case 'crear': {
        if (st.ocupado || st.nombre.trim().length < 3) return;
        st.ocupado = true; render();
        try {
          const nd = await crearDespacho(st.nombre);
          st.ocupado = false;
          st.prev = 'home';
          vista = { id: nd.id, nube: [], cargando: false, error: null };
          ir('despacho');
          toast('Carpeta creada en OneDrive');
        } catch (err) {
          st.ocupado = false; render();
          toast(errorAMensaje(err, 'crear la carpeta'));
        }
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
        st.revision = null; // las URLs se liberan cuando la foto termina de subirse
        items.forEach(it => URL.revokeObjectURL(it.url));
        const n = await encolar(items, d);
        render();
        toast(items.length === 1 ? `Foto ${pad3(n)} aceptada` : `${items.length} fotos aceptadas`);
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
        st.ocupado = true; render();
        try {
          const n = await contarFotos(d.id);
          await cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); x.estado = 'cerrado'; x.cerrado = new Date().toISOString(); x.fotos = n; });
          st.ocupado = false;
          ir('home');
          toast(`Despacho "${d.nombre}" cerrado`);
        } catch (err) {
          st.ocupado = false; render();
          toast(errorAMensaje(err, 'cerrar el despacho'));
        }
        break;
      }
      case 'reabrir': {
        st.ocupado = true; render();
        try {
          await cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); x.estado = 'proceso'; x.cerrado = null; });
          toast('Despacho reabierto');
        } catch (err) { toast(errorAMensaje(err, 'reabrir el despacho')); }
        st.ocupado = false; render();
        break;
      }
      case 'estado': st.estado = b.dataset.e; render(); break;
      case 'limpiar': Object.assign(st, { q: '', desde: '', hasta: '', estado: 'todos' }); render(); break;
    }
  });

  document.addEventListener('input', e => {
    const t = e.target;
    if (t.id === 'nombre') {
      st.nombre = t.value;
      const { nombre, quitados } = nombreCarpeta(st.nombre, new Date());
      document.getElementById('ruta').innerHTML = vistaRuta(nombre, quitados);
      document.getElementById('crear').disabled = st.ocupado || st.nombre.trim().length < 3;
    } else if (t.id === 'q' || t.id === 'desde' || t.id === 'hasta') {
      st[t.id] = t.value;
      document.getElementById('lista').innerHTML = listaRegistro();
    }
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.target.id === 'nombre') document.getElementById('crear').click();
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
    ir('home');
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

    registro = cache.leer('registro.' + cuenta.username) || { despachos: [] };
    await refrescarCola();
    ir('home');
    try { await cargarRegistro(); refrescar(); }
    catch (e) { if (!(e instanceof SinSenal)) { console.error(e); toast('No se pudo actualizar el registro desde OneDrive.'); } }
    procesarCola();
  }

  if (!window.HUB_DEMO && 'serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  arrancar();
})();
