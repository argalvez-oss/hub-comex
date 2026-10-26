/* Hub Comex — app de Ventas para registrar despachos y traslados con fotos y generar su informe en Word.
 *
 * Modos (config.js → modo):
 *   'firebase'  : registros y fotos en Firebase (Firestore), sincronizados entre dispositivos y cuentas. Provisional.
 *   'microsoft' : ingreso con Microsoft y todo en OneDrive (a futuro, SharePoint).
 *   'prueba'    : sin nube; todo queda en el dispositivo.
 * - Ingreso: Firebase (correo y contraseña), Microsoft (MSAL) o modo prueba (se elige Ventas / Soporte). La sesión queda guardada.
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
  // Modo prueba (config modo: 'prueba', o …/#demo): sin nube. Se elige Ventas o Soporte al ingresar
  // y los registros y fotos se guardan solo en este dispositivo.
  const DEMO = !['microsoft', 'firebase'].includes(CFG.modo) || location.hash === '#demo' || window.HUB_DEMO === true;
  const FB = !DEMO && CFG.modo === 'firebase';
  const SIN_ONEDRIVE = DEMO || FB; // sin carpetas de OneDrive: se ocultan los botones y rutas de carpeta
  // Con cuentas (firebaseCuentas: true) se ingresa con correo y contraseña. Sin cuentas, se elige Ventas o Soporte
  // y la app se identifica ante Firebase de forma anónima (invisible), para que la base solo acepte conexiones de la app.
  const FB_CUENTAS = FB && CFG.firebaseCuentas === true;
  const FB_SDK = ['app', 'auth', 'firestore'].map(m => `https://www.gstatic.com/firebasejs/10.14.1/firebase-${m}-compat.js`);
  const ELEGIR_CUENTA = DEMO || (FB && !FB_CUENTAS); // ingreso eligiendo Ventas / Soporte
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
  // Campos del informe, por sección, en el mismo orden del formato. Cada campo: [clave, etiqueta, ejemplo, opciones].
  //   lista: nombre del catálogo de config.js para la lista desplegable · opcional: no cuenta como faltante al cerrar
  //   num: teclado numérico · opciones: botones para elegir (ej. Reefer / Camión)
  const SECCIONES = [
    { id: 'generales', titulo: 'Datos generales', campos: [
      ['cliente', 'Cliente', 'Ej. Goya Foods Inc.', { lista: 'clientes' }],
      ['destino', 'Destino', 'Ej. Jersey City, EE. UU.', { lista: 'destinos' }],
      ['lugarCarga', 'Lugar de carga', 'Ej. Planta Quevedo', { lista: 'lugaresCarga' }],
      ['pedido', 'Pedido / OC', 'Ej. 267-261523'],
      ['factura', 'Factura / DAE', 'Ej. 001-001-000123'],
      ['responsable', 'Responsable', 'Nombre de quien despacha'],
      ['ci', 'C.I.', 'Cédula del responsable', { num: true }]
    ] },
    { id: 'transporte', titulo: 'Transporte y contenedor', campos: [
      ['transportista', 'Transportista', 'Empresa de transporte', { lista: 'transportistas' }],
      ['conductor', 'Conductor', 'Nombre del conductor'],
      ['cedula', 'Cédula', 'Cédula del conductor', { num: true }],
      ['placa', 'Placa', 'Ej. GBA-1234'],
      ['contenedor', 'Contenedor', 'Ej. MSKU 123456-7'],
      ['tipoUnidad', 'Tipo', '', { opciones: [['reefer', 'Reefer'], ['camion', 'Camión']] }],
      ['naviera', 'Naviera', 'Ej. Maersk', { lista: 'navieras' }],
      ['capacidad', 'Capacidad', "Ej. 40'"],
      ['tInicial', 'T. inicial (°C)', 'Ej. -18'],
      ['tFinal', 'T. final (°C)', 'Ej. -19'],
      ['setPoint', 'Set point (°C)', 'Ej. -20'],
      ['termografo', 'Termógrafo', 'N.º del termógrafo']
    ] },
    { id: 'insp', titulo: 'Inspección previa de la unidad' },
    { id: 'sellos', titulo: 'Estiba, sellos y cierre', campos: [
      ['selloNaviera', 'Sello naviera', ''], ['selloBotella', 'Sello botella', ''],
      ['selloCable', 'Sello cable', ''], ['sticker', 'Sticker', ''],
      ['otrosSellos', 'Otros sellos', '', { opcional: true }]
    ] },
    { id: 'carga', titulo: 'Detalle de carga' },
    { id: 'obs', titulo: 'Observaciones' },
    { id: 'personal', titulo: 'Personal a cargo', campos: [
      ['jefeCalidad', 'Jefe de Calidad', 'Nombre'],
      ['respSellos', 'Responsable de sellos', 'Nombre'],
      ['bodega1', 'Bodega', 'Nombre'],
      ['bodega2', 'Bodega (2)', 'Nombre', { opcional: true }],
      ['bodega3', 'Bodega (3)', 'Nombre', { opcional: true }]
    ] }
  ];
  const CAMPOS = SECCIONES.flatMap(s => s.campos || []);
  const CAMPO = Object.fromEntries(CAMPOS.map(c => [c[0], c]));
  const NUEVO_CAMPOS = ['cliente', 'destino', 'lugarCarga', 'pedido'].map(k => CAMPO[k]); // lo básico al crear
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
    spin: '<path d="M12 3a9 9 0 1 0 9 9"/>',
    trash: '<path d="M4 7h16M9 7V4.5h6V7M6.5 7l1 13h9l1-13M10 11v6M14 11v6"/>',
    star: '<path d="m12 3.5 2.6 5.3 5.9.9-4.25 4.1 1 5.85L12 16.9l-5.25 2.75 1-5.85L3.5 9.7l5.9-.9z"/>'
  };
  const ic = n => `<svg class="i" viewBox="0 0 24 24" aria-hidden="true">${ICONS[n]}</svg>`;
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pad = n => String(n).padStart(2, '0');
  const pad3 = n => String(n).padStart(3, '0');
  const isoDay = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fmtDate = d => new Date(d).toLocaleDateString('es-EC', { day: 'numeric', month: 'short', year: 'numeric' });
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
    if (FB) return fbGraph(ruta, op);
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
  // Sin OneDrive: arranca vacío, solo con las carpetas de Despachos y Traslados.
  function demoSemilla() {
    const raiz = demoCarpeta(RAIZ, null).id;
    demoCarpeta('Despachos', raiz);
    demoCarpeta('Traslados', raiz);
    demo.registro = JSON.stringify({ despachos: [] });
  }
  // Lo simulado se guarda en este dispositivo: datos en localStorage, fotos en IndexedDB.
  const demoClave = () => 'demo.datos.v6'; // compartido entre Ventas y Soporte en este dispositivo
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
      const carpeta = demo.carpetas.get(m[1]);
      if (carpeta) { // eliminar un registro completo con sus fotos
        for (const f of carpeta.fotos) await idb.guardar('demofoto.' + f.id, null);
        demo.carpetas.delete(m[1]);
        demoGuardar();
        return null;
      }
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

  // ---------- Firebase (provisional, hasta pasar a SharePoint) ----------
  // Firestore: "registros" (un documento por despacho o traslado), "fotos" (foto reducida a calidad de informe)
  // y "miniaturas" (versión pequeña para mostrar en pantalla). Sin señal, Firestore guarda en el dispositivo y sube luego.
  let fbAuth = null, fbDb = null, fbEscucha = null;
  const cargarScript = src => new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src; s.onload = res; s.onerror = () => rej(new SinSenal());
    document.head.appendChild(s);
  });
  async function iniciarFirebase() {
    for (const src of FB_SDK) await cargarScript(src);
    firebase.initializeApp(CFG.firebase);
    fbDb = firebase.firestore();
    try { await fbDb.enablePersistence({ synchronizeTabs: true }); } catch (_) { /* el navegador no lo permite: funciona igual, solo con señal */ }
    fbAuth = firebase.auth();
    const u = await new Promise(res => { const off = fbAuth.onAuthStateChanged(x => { off(); res(x); }); });
    if (FB_CUENTAS) return u;
    if (!u) await anonimo(); // queda guardado en el dispositivo; solo se hace la primera vez
    return null;
  }
  // Identificación anónima ante Firebase. Sin señal se reintenta al volver la conexión.
  async function anonimo() {
    try { await fbAuth.signInAnonymously(); }
    catch (e) { console.warn('Firebase anónimo', e.code); window.addEventListener('online', () => { if (!fbAuth.currentUser) anonimo(); }, { once: true }); }
  }
  const limpio = o => JSON.parse(JSON.stringify(o)); // Firestore no acepta valores "undefined"
  const bytesFb = u8 => firebase.firestore.Blob.fromUint8Array(u8);
  const blobDe = b => new Blob([b.toUint8Array()], { type: 'image/jpeg' });
  // Espera la confirmación del servidor sin quedarse colgado: sin señal, la cola reintenta luego con el mismo id.
  const conPlazo = (p, ms = 25000) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new SinSenal()), ms))]);

  // JPEG dentro de un tamaño máximo, para que quepa en un documento de Firestore (límite 1 MB).
  async function jpeg(blob, lado, calidad, maxBytes) {
    const img = await createImageBitmap(blob);
    let k = Math.min(1, lado / Math.max(img.width, img.height)), q = calidad, b, c;
    for (let i = 0; i < 6; i++) {
      c = document.createElement('canvas');
      c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      b = await new Promise(r => c.toBlob(r, 'image/jpeg', q));
      if (b.size <= maxBytes) break;
      q = Math.max(0.45, q - 0.1); k *= 0.85;
    }
    return { u8: new Uint8Array(await b.arrayBuffer()), w: c.width, h: c.height };
  }

  async function fbLeerRegistros() {
    const snap = await fbDb.collection('registros').get();
    return { despachos: snap.docs.map(x => x.data()) };
  }
  // Guarda solo los registros que cambiaron (con quién y cuándo). Sin señal, Firestore los deja en cola.
  function fbGuardarCambios(antes, data) {
    const previo = new Map(antes.despachos.map(x => [x.id, JSON.stringify(x)]));
    const lote = fbDb.batch();
    const ahora = new Date().toISOString();
    let n = 0;
    for (const x of data.despachos) {
      const igual = previo.get(x.id) === JSON.stringify(x);
      previo.delete(x.id);
      if (igual) continue;
      x.modificado = ahora; x.modificadoPor = cuenta.username;
      lote.set(fbDb.collection('registros').doc(x.id), limpio(x)); n++;
    }
    for (const id of previo.keys()) { lote.delete(fbDb.collection('registros').doc(id)); n++; }
    if (n) lote.commit().catch(e => { console.error('Firestore', e); toast('No se pudo guardar en la nube. Avisa a soporte.'); });
  }
  // Tiempo real: lo que se registra en otro dispositivo o cuenta aparece aquí sin recargar.
  function fbEscuchar() {
    if (fbEscucha) fbEscucha();
    fbEscucha = fbDb.collection('registros').onSnapshot(snap => {
      registro = aplicarPendientes({ despachos: snap.docs.map(x => x.data()) });
      cache.guardar('registro.' + usuarioDatos(), registro);
      refrescar();
      const lista = document.getElementById('lista');
      if (st.view === 'registros' && lista && document.activeElement?.id !== 'q') lista.innerHTML = listaRegistro();
    }, e => {
      console.error('Firestore', e);
      // Si aún no hay identificación (p. ej. se abrió sin señal), se vuelve a escuchar cuando la haya.
      if (e.code === 'permission-denied' && fbAuth) { const off = fbAuth.onAuthStateChanged(u => { if (u) { off(); fbEscuchar(); } }); }
    });
  }
  // Las mismas operaciones que usa la app con OneDrive, pero sobre Firestore.
  async function fbGraph(ruta, op) {
    const metodo = op.method || 'GET';
    let m;
    if ((m = /^\/me\/drive\/items\/([^/:?]+)\/children\?/.exec(ruta))) { // fotos de un registro (miniaturas)
      const snap = await fbDb.collection('miniaturas').where('registroId', '==', m[1]).get();
      return { value: snap.docs.map(x => { const f = x.data(); return { id: x.id, name: f.nombre, file: {}, thumbnails: f.data ? [{ medium: { url: URL.createObjectURL(blobDe(f.data)) } }] : [] }; }) };
    }
    if ((m = /^\/me\/drive\/items\/([^/:?]+)$/.exec(ruta)) && metodo === 'DELETE') {
      const id = m[1];
      const lote = fbDb.batch();
      if (buscar(id)) { // registro completo: se borran todas sus fotos
        const snap = await fbDb.collection('miniaturas').where('registroId', '==', id).get();
        snap.docs.forEach(x => { lote.delete(x.ref); lote.delete(fbDb.collection('fotos').doc(x.id)); });
      } else {
        lote.delete(fbDb.collection('fotos').doc(id));
        lote.delete(fbDb.collection('miniaturas').doc(id));
      }
      await conPlazo(lote.commit());
      return null;
    }
    throw new Error('Firebase: operación no soportada ' + ruta);
  }
  async function fbSubirFoto(it) {
    if (!navigator.onLine) throw new SinSenal();
    // Foto grande lo más nítida posible dentro del límite de Firestore (1 MB por documento); miniatura solo para la lista.
    const [grande, mini] = await Promise.all([jpeg(it.blob, 1920, 0.86, 950 * 1024), jpeg(it.blob, 360, 0.6, 60 * 1024)]);
    const comun = { registroId: it.despachoId, nombre: it.archivo, creado: new Date().toISOString(), creadoPor: cuenta.username };
    const lote = fbDb.batch();
    lote.set(fbDb.collection('fotos').doc(it.id), { ...comun, data: bytesFb(grande.u8), w: grande.w, h: grande.h });
    lote.set(fbDb.collection('miniaturas').doc(it.id), { ...comun, data: bytesFb(mini.u8) });
    await conPlazo(lote.commit(), 60000);
    return { id: it.id, name: it.archivo };
  }
  async function fbFotoGrande(id) {
    const x = await fbDb.collection('fotos').doc(id).get();
    if (!x.exists) throw new Error('no está');
    return blobDe(x.data().data);
  }

  // ---------- Registro (despachos y traslados) ----------
  let registro = { despachos: [] };
  // En modo prueba Ventas y Soporte comparten los datos del dispositivo; con Microsoft, cada cuenta usa su OneDrive.
  // En Firebase todo está en la nube compartida, así que lo local también se comparte entre cuentas del dispositivo.
  const usuarioDatos = () => (DEMO || FB ? 'compartido' : cuenta.username);
  const clavePend = () => 'pendientes.' + usuarioDatos();

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
    if (FB) {
      data = await fbLeerRegistros();
      data.despachos.forEach(x => { if (!x.tipo) x.tipo = 'despacho'; });
      return data;
    }
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
    cache.guardar('registro.' + usuarioDatos(), registro);
    return registro;
  }
  // Lee lo último de OneDrive, aplica los cambios pendientes y el cambio pedido, y lo guarda.
  async function cambiarRegistro(cambio) {
    const pendAntes = JSON.stringify(cache.leer(clavePend()) || {});
    const base = await leerRegistro();
    const antes = FB ? JSON.parse(JSON.stringify(base)) : null;
    const data = aplicarPendientes(base);
    cambio(data);
    if (FB) fbGuardarCambios(antes, data);
    else await graph(REGISTRO + ':/content', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data, null, 1) });
    if (JSON.stringify(cache.leer(clavePend()) || {}) === pendAntes) cache.quitar(clavePend());
    registro = aplicarPendientes(data);
    cache.guardar('registro.' + usuarioDatos(), registro);
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
    cache.guardar('registro.' + usuarioDatos(), registro);
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
    let f;
    if (FB) f = { id: fbDb.collection('registros').doc().id, name: nombre, webUrl: '' }; // en Firebase no hay carpetas
    else {
      const padre = await carpetaTipo(tipo);
      // "rename": si ya existe una carpeta con ese nombre, OneDrive agrega un número al final
      f = await graph(`/me/drive/items/${padre}/children`, { method: 'POST', json: { name: nombre, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' } });
    }
    const d = {
      id: f.id, tipo, nombre: datos.nombre.trim().replace(/\s+/g, ' '), carpeta: f.name, webUrl: f.webUrl,
      fecha: fecha.toISOString(), estado: 'proceso', cerrado: null, fotos: 0, inspeccion: {},
      productos: [], totalCajas: '', totalPallets: '', observaciones: '', informeFotos: [],
      creado: fecha.toISOString(), creadoPor: cuenta.username // trazabilidad
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
    cola = (await idb.cola()).filter(x => x.cuenta === usuarioDatos()).sort((a, b) => a.creado - b.creado);
    for (const it of cola) if (!urlsCola.has(it.id)) urlsCola.set(it.id, URL.createObjectURL(it.blob));
  }
  const pendientesDe = id => cola.filter(x => x.despachoId === id);

  async function subirArchivo(it) {
    if (FB) return fbSubirFoto(it);
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
        id: uid(), cuenta: usuarioDatos(), despachoId: d.id, folderId: d.id, despacho: d.nombre,
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
  // Reduce una imagen para el informe (1600 px, JPEG) y devuelve sus bytes y medidas.
  async function prepararFoto(blob) {
    const img = await createImageBitmap(blob);
    const k = Math.min(1, 1600 / Math.max(img.width, img.height));
    const c = document.createElement('canvas');
    c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    const b = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.86));
    return { data: new Uint8Array(await b.arrayBuffer()), w: c.width, h: c.height };
  }
  // Foto en tamaño completo (no la miniatura): de la nube, de OneDrive o del dispositivo, según el modo.
  async function fotoNube(id) {
    if (DEMO) { const b = await idb.leer('demofoto.' + id); if (!b) throw new Error('sin archivo'); return b; }
    if (FB) return fbFotoGrande(id);
    const meta = await graph(`/me/drive/items/${id}`);
    const resp = await fetch(meta['@microsoft.graph.downloadUrl']);
    if (!resp.ok) throw new Error('descarga');
    return resp.blob();
  }
  function bajarArchivo(blob, nombre) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = nombre.replace(/[\\/:*?"<>|]/g, '-');
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 10000);
  }

  // ---------- Descargar todas las fotos de un registro (.zip) ----------
  const JSZIP_URL = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
  function cargarJSZip() {
    if (window.JSZip) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = JSZIP_URL;
      s.onload = () => (window.JSZip ? res() : rej(new Error('JSZip no disponible')));
      s.onerror = () => { s.remove(); rej(new SinSenal()); };
      document.head.appendChild(s);
    });
  }
  async function descargarZip(d) {
    if (window.HUB_DEMO) { toast('Desde este enlace no se pueden descargar archivos. Usa la app publicada en GitHub.'); return; }
    st.generandoZip = d.id; render();
    try {
      await cargarJSZip();
      const archivos = new Map(); // nombre → foto
      for (const x of pendientesDe(d.id)) archivos.set(x.archivo, x.blob); // las que aún están en el celular
      let faltan = 0;
      const r = await graph(`/me/drive/items/${d.id}/children?$select=id,name,file&$top=999`);
      const enNube = r.value.filter(f => f.file && ES_FOTO.test(f.name) && !archivos.has(f.name));
      for (const f of enNube) {
        try { archivos.set(f.name, await fotoNube(f.id)); } catch (_) { faltan++; }
      }
      if (!archivos.size) { toast('Este registro todavía no tiene fotos.'); return; }
      const zip = new window.JSZip();
      [...archivos.keys()].sort((a, b) => a.localeCompare(b, 'es', { numeric: true }))
        .forEach(n => zip.file(n, archivos.get(n), { binary: true }));
      const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' }); // las fotos ya vienen comprimidas
      bajarArchivo(blob, `${d.carpeta || d.nombre}.zip`);
      toast(faltan ? `Se descargaron ${archivos.size} fotos; ${faltan} no se pudieron bajar (revisa la señal).` : `Se descargaron ${archivos.size} fotos.`);
    } catch (e) {
      console.error(e);
      toast(errorAMensaje(e, 'descargar las fotos'));
    } finally {
      st.generandoZip = null; render();
    }
  }
  // Fotos marcadas con estrella, en orden: se toman del celular si aún no se suben, o de OneDrive.
  async function fotosParaInforme(d) {
    const marcadas = (d.informeFotos || []).slice().sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
    const blobs = new Map();
    for (const x of pendientesDe(d.id)) if (marcadas.includes(x.archivo)) blobs.set(x.archivo, x.blob);
    let faltan = 0;
    const restantes = marcadas.filter(n => !blobs.has(n));
    if (restantes.length) {
      try {
        const r = await graph(`/me/drive/items/${d.id}/children?$select=id,name,file&$top=999`);
        for (const n of restantes) {
          const it = r.value.find(f => f.name === n);
          try {
            if (!it) throw new Error('no está');
            blobs.set(n, await fotoNube(it.id));
          } catch (_) { faltan++; }
        }
      } catch (_) { faltan += restantes.length; }
    }
    const fotos = [];
    for (const n of marcadas) {
      if (!blobs.has(n)) continue;
      try { fotos.push({ nombre: n, ...(await prepararFoto(blobs.get(n))) }); } catch (_) { faltan++; }
    }
    return { fotos, faltan };
  }
  async function descargarInforme(d) {
    if (window.HUB_DEMO) { toast('Desde este enlace no se pueden descargar archivos. Usa la app publicada en GitHub.'); return; }
    st.generando = d.id; render();
    try {
      await cargarLibreriaWord();
      const { fotos, faltan } = await fotosParaInforme(d); // el logo ya viene dentro de informe.js
      if (faltan) toast(`${faltan} ${faltan === 1 ? 'foto marcada no se pudo agregar' : 'fotos marcadas no se pudieron agregar'} (sin señal o ya no están). Puedes pegarlas en Word.`);
      const doc = window.HCInforme.crear(window.docx, d, { ...(CFG.informe || {}), fotos });
      const blob = await window.docx.Packer.toBlob(doc);
      bajarArchivo(blob, window.HCInforme.nombreArchivo(d));
      setTimeout(() => toast('Informe descargado. Ábrelo en Word para revisarlo.'), faltan ? 3800 : 0);
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
    nuevo: {}, q: '', desde: '', hasta: '', estado: 'todos', error: null, abiertos: {}, borrar: null, generando: null,
    libre: {}, // campos de lista donde se eligió "Otro (escribir)"
    sel: null, // lista desplegable abierta
    eliminar: null // registro que se está por eliminar (pide confirmación)
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
    if (a && app.contains(a) && /INPUT|TEXTAREA|SELECT/.test(a.tagName)) { repintar = true; return; }
    render();
  }
  document.addEventListener('focusout', () => {
    if (!repintar) return;
    setTimeout(() => {
      const a = document.activeElement;
      if (a && app.contains(a) && /INPUT|TEXTAREA|SELECT/.test(a.tagName)) return;
      repintar = false; refrescar();
    }, 250);
  });
  const saludo = () => { const h = new Date().getHours(); return h < 12 ? 'Buenos días' : h < 19 ? 'Buenas tardes' : 'Buenas noches'; };
  const ordenados = tipo => registro.despachos.filter(d => !tipo || d.tipo === tipo).sort((a, b) => b.fecha.localeCompare(a.fecha));
  const inspHechos = d => INSPECCION.filter(([k]) => d.inspeccion && d.inspeccion[k] && d.inspeccion[k].v).length;
  // Avance de cada sección: [llenos, total], sin contar los campos opcionales.
  function avance(d, s) {
    if (s.id === 'insp') return [inspHechos(d), INSPECCION.length];
    if (s.id === 'carga') {
      const hay = (d.productos || []).some(p => (p.producto || '').trim());
      return [(hay ? 1 : 0) + ['totalCajas', 'totalPallets'].filter(k => (d[k] || '').trim()).length, 3];
    }
    if (s.id === 'obs') return [(d.observaciones || '').trim() ? 1 : 0, 0];
    const req = s.campos.filter(c => !(c[3] && c[3].opcional));
    return [req.filter(([k]) => (d[k] || '').trim()).length, req.length];
  }
  const faltantes = d => SECCIONES.map(s => ({ s, a: avance(d, s) })).filter(x => x.a[1] && x.a[0] < x.a[1]);
  const delMes = d => { const f = new Date(d.fecha), h = new Date(); return f.getMonth() === h.getMonth() && f.getFullYear() === h.getFullYear(); };

  // ---------- Pantallas ----------
  function vIngreso() {
    if (FB_CUENTAS) {
      return `<main class="wrap"><section class="login">
        <div class="mark">HC</div>
        <h1>Hub Comex</h1>
        <form id="f-login" class="login-form" novalidate>
          <div class="field"><label for="l-correo">Correo</label>
            <input class="input" id="l-correo" type="email" autocomplete="username" inputmode="email" placeholder="ventas-sales@healthyfood.com.ec" value="${esc(st.correo || '')}"></div>
          <div class="field"><label for="l-clave">Contraseña</label>
            <input class="input" id="l-clave" type="password" autocomplete="current-password"></div>
          ${st.errorLogin ? `<p class="warn">${ic('alert')}<span>${esc(st.errorLogin)}</span></p>` : ''}
          <button class="btn primary big block" type="submit" ${st.ocupado ? 'disabled' : ''}>${st.ocupado ? ic('spin') + 'Ingresando…' : ic('lock') + 'Iniciar sesión'}</button>
          <button type="button" class="link" data-act="olvide">¿Olvidaste tu contraseña?</button>
        </form>
      </section></main>`;
    }
    return `<main class="wrap"><section class="login">
      <div class="mark">HC</div>
      <h1>Hub Comex</h1>
      <button class="btn primary big block" data-act="ingresar">${ic('lock')}Iniciar sesión</button>
      ${ELEGIR_CUENTA && st.opcionesDemo ? `<div class="demo-box">
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
        ? `<button class="areas-btn" data-act="areas" aria-label="Volver al inicio">${ic('back')}<span>Inicio</span></button>
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
      <button class="nav-b" data-act="ir" data-v="modulo" ${v === 'modulo' ? 'aria-current="page"' : ''}>${ic(st.tipo === 'traslado' ? 'swap' : 'truck')}${t.plural}</button>
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
      ${SIN_ONEDRIVE ? '' : `<a class="icon-btn" href="${esc(d.webUrl)}" target="_blank" rel="noopener" aria-label="Abrir carpeta en OneDrive" title="Abrir carpeta">${ic('folder')}</a>`}
    </div>`;
  }
  function notas() {
    const pend = enCola().length;
    return `${st.rol === 'soporte' && !SIN_ONEDRIVE ? `<div class="note info">${ic('info')}<span>Estás en modo soporte. Tus pruebas se guardan en tu propio OneDrive, no en el de Ventas.</span></div>` : ''}
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
      <div class="areas-head"><h1>${saludo()}, ${nombre}</h1><p>¿Qué vas a registrar hoy?</p></div>
      ${notas()}
      ${st.migrar ? `<div class="note">${ic('cloud')}<span>Hay <b>${st.migrar} ${st.migrar === 1 ? 'registro guardado' : 'registros guardados'} solo en este dispositivo</b>, de antes de conectar la nube.
        <button class="link" data-act="migrar" ${st.ocupado ? 'disabled' : ''}>${st.ocupado ? 'Subiendo…' : 'Subirlos a la nube'}</button></span></div>` : ''}
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
      if (q && ![d.nombre, d.carpeta, d.cliente, d.destino, d.pedido, d.contenedor, d.factura].some(x => (x || '').toLowerCase().includes(q))) return false;
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

  // ---------- Listas desplegables ----------
  // Las opciones salen de catalogos.js (generado desde el Excel "LISTA DESPLEGABLE APP COMEX.xlsx").
  // Los campos con lista se eligen de un desplegable; "Otro (escribir)" permite poner un valor que no esté.
  // Los productos son estrictos: solo los del cliente elegido, y al elegir uno se llena su código.
  const norm = s => (s || '').toString().trim().toLowerCase();
  const catalogo = () => window.HUB_CATALOGOS || CFG.catalogos || {};
  const OTRO = '__otro';
  function opcionesLista([k, , , o]) {
    const cat = (o && o.lista && catalogo()[o.lista]) || [];
    const hist = registro.despachos.map(d => (d[k] || '').trim()).filter(v => v && !cat.some(c => norm(c) === norm(v)));
    return [...new Set([...cat, ...hist.sort()])];
  }
  // Productos del catálogo para un cliente (lista estricta).
  const productosCatalogo = cliente => (catalogo().productos || []).filter(p => norm(p.cliente) === norm(cliente));
  function codigoDe(cliente, producto) {
    const enCat = productosCatalogo(cliente).find(p => norm(p.producto) === norm(producto));
    if (enCat) return enCat.codigo || '';
    const usado = registro.despachos.filter(d => norm(d.cliente) === norm(cliente)).flatMap(d => d.productos || []).find(p => norm(p.producto) === norm(producto) && p.codigo);
    return usado ? usado.codigo : '';
  }
  // Sugerencias para los campos de texto libre (lo ya escrito en registros anteriores).
  function sugerencias() {
    return CAMPOS.filter(c => !(c[3] && (c[3].opciones || c[3].lista))).map(([k]) => {
      const vals = [...new Set(registro.despachos.map(d => (d[k] || '').trim()).filter(Boolean))].slice(0, 100);
      return vals.length ? `<datalist id="dl-${k}">${vals.map(v => `<option value="${esc(v)}"></option>`).join('')}</datalist>` : '';
    }).join('');
  }
  // Lista desplegable propia (en lugar de la del navegador), con el mismo estilo de la app.
  function selector(id, valor, opciones, placeholder) {
    const abierto = st.sel === id;
    const elegido = opciones.find(o => !o.otro && norm(o.v) === norm(valor));
    const texto = elegido ? elegido.l : valor || '';
    return `<div class="sel ${abierto ? 'open' : ''}">
      <button type="button" class="input sel-btn" id="${id}" data-act="abrirsel" data-id="${id}" aria-haspopup="listbox" aria-expanded="${abierto}">
        <span class="${texto ? '' : 'ph'}">${esc(texto || placeholder)}</span>${ic('chev')}
      </button>
      ${abierto ? `<div class="sel-pop" role="listbox">${opciones.map(o => {
        const activo = !o.otro && norm(o.v) === norm(valor);
        return `<button type="button" class="sel-op ${o.otro ? 'otro' : ''}" role="option" aria-selected="${activo}" data-act="elegir" data-id="${id}" data-v="${esc(o.v)}"><span>${esc(o.l)}</span>${activo ? ic('check') : ''}</button>`;
      }).join('')}</div>` : ''}
    </div>`;
  }
  function campo(def, valor, prefijo) {
    const [k, etiqueta, ej, o = {}] = def;
    const id = `${prefijo}-${k}`;
    if (o.opciones) {
      return `<div class="field"><span class="lbl">${etiqueta}</span><div class="seg" role="group" aria-label="${esc(etiqueta)}">
        ${o.opciones.map(([v, l]) => `<button data-act="opcion" data-k="${k}" data-v="${v}" aria-pressed="${valor === v}">${l}</button>`).join('')}</div></div>`;
    }
    const opciones = o.lista ? opcionesLista(def) : [];
    if (o.lista && opciones.length && !st.libre[id]) {
      return `<div class="field"><label for="${id}">${etiqueta}</label>
        ${selector(id, valor, [...opciones.map(v => ({ v, l: v })), { v: OTRO, l: 'Otro (escribir)', otro: true }], 'Elegir…')}</div>`;
    }
    return `<div class="field"><label for="${id}">${etiqueta}${o.lista && opciones.length ? ` <button class="link mini" data-act="alista" data-id="${id}">Elegir de la lista</button>` : ''}</label>
      <input class="input" id="${id}" data-campo="${k}" ${o.lista ? '' : `list="dl-${k}"`} autocomplete="off" ${o.num ? 'inputmode="numeric"' : ''} placeholder="${esc(ej)}" value="${esc(valor || '')}" maxlength="120"></div>`;
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
      <div class="preview-path" id="ruta" ${SIN_ONEDRIVE ? 'hidden' : ''}>${vistaRuta(nombre, quitados)}</div>
      <div class="form-grid">${NUEVO_CAMPOS.map(c => campo(c, n[c[0]], 'nuevo')).join('')}</div>
      <p class="help">La fecha se guarda sola al crear. El resto de datos (transporte, sellos, productos, personal) se llena dentro del ${t.uno}.</p>
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
    else if (!esperando) { cls = 'ok'; icon = 'cloudok'; linea = `${total} ${total === 1 ? 'foto subida' : 'fotos subidas'}`; ayuda = DEMO ? 'Están guardadas en la carpeta.' : FB ? 'Ya están en la nube.' : 'Ya están en OneDrive.'; }
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

  // ---------- Secciones del registro (mismo orden que el informe) ----------
  const tituloSec = (s, d) => s.id === 'personal' ? `Personal a cargo del ${T(d).uno}` : s.titulo;
  function plegable(d, s, i, contenido) {
    const [hechos, total] = avance(d, s);
    const badge = total ? `<span class="prog ${hechos >= total ? 'ok' : ''}">${hechos} de ${total}</span>` : `<span class="prog ${hechos ? 'ok' : 'opc'}">${hechos ? 'Escrito' : 'Opcional'}</span>`;
    return `<details class="sec" id="sec-${s.id}" data-sec="${s.id}" ${st.abiertos[s.id] ? 'open' : ''}>
      <summary><span><small class="sec-num">${i + 1}</small>${esc(tituloSec(s, d))}</span>${badge}${ic('chev')}</summary>
      <div class="sec-in">${contenido}</div>
    </details>`;
  }
  function seccionCampos(d, s) {
    return `<div class="form-grid">${s.campos.map(c => campo(c, d[c[0]], 'det')).join('')}</div>
      ${s.id === 'generales' ? `<div class="kv"><span>Fecha</span><b>${fmtDate(d.fecha)}</b></div>` : ''}`;
  }
  function seccionInspeccion(d) {
    const insp = d.inspeccion || {};
    return `<div class="insp">${INSPECCION.map(([k, nombre]) => {
      const r = insp[k] || {};
      return `<div class="insp-row">
        <span class="insp-n">${esc(nombre)}</span>
        <div class="seg" role="group" aria-label="${esc(nombre)}">
          <button data-act="insp" data-k="${k}" data-v="si" aria-pressed="${r.v === 'si'}">Sí</button>
          <button data-act="insp" data-k="${k}" data-v="no" aria-pressed="${r.v === 'no'}">No</button>
        </div>
        ${r.v === 'no' ? `<input class="input obs" data-obs="${k}" placeholder="¿Qué se encontró?" value="${esc(r.obs || '')}" maxlength="200" aria-label="Observación: ${esc(nombre)}">` : ''}
      </div>`;
    }).join('')}</div>
    <p class="help">Se marca antes de cargar. Si algo está mal, escribe la observación; sale en el informe.</p>`;
  }
  function seccionCarga(d) {
    const prods = d.productos && d.productos.length ? d.productos : [{}];
    const lista = productosCatalogo(d.cliente);
    // Con un cliente del catálogo: solo se puede elegir entre sus productos. Sin cliente del catálogo: se escribe.
    const selProducto = (p, i) => {
      const fuera = p.producto && !lista.some(x => norm(x.producto) === norm(p.producto));
      return `<div class="p-prod">${selector(`prod-${i}`, fuera ? `${p.producto} (no es de este cliente)` : p.producto,
        lista.map(x => ({ v: x.producto, l: x.producto })), 'Elegir producto…')}</div>`;
    };
    return `${lista.length ? `<p class="help">Productos de <b>${esc(d.cliente)}</b>. Al elegir uno, el código se llena solo.</p>`
        : d.cliente ? `<p class="help"><b>${esc(d.cliente)}</b> no tiene productos en la lista; escríbelos.</p>`
        : `<p class="help">Elige primero el <b>cliente</b> en Datos generales para ver solo sus productos.</p>`}
      <div class="prods">
        ${prods.map((p, i) => `<div class="prod">
          <span class="prod-n">${i + 1}</span>
          ${lista.length ? selProducto(p, i) : `<input class="input p-prod" data-prod="producto" data-i="${i}" autocomplete="off" placeholder="Producto" value="${esc(p.producto || '')}" aria-label="Producto ${i + 1}">`}
          <input class="input p-cod" data-prod="codigo" data-i="${i}" autocomplete="off" placeholder="Código" value="${esc(p.codigo || '')}" aria-label="Código ${i + 1}">
          <input class="input p-can" data-prod="cantidad" data-i="${i}" autocomplete="off" inputmode="decimal" placeholder="Cantidad" value="${esc(p.cantidad || '')}" aria-label="Cantidad ${i + 1}">
          <button class="icon-btn p-x" data-act="quitarprod" data-i="${i}" aria-label="Quitar producto ${i + 1}">${ic('x')}</button>
        </div>`).join('')}
      </div>
      <button class="btn" data-act="agregarprod">${ic('plus')}Agregar producto</button>
      <div class="form-grid">${campo(['totalCajas', 'Total cajas', 'Ej. 2000', { num: true }], d.totalCajas, 'det')}${campo(['totalPallets', 'Total pallets', 'Ej. 20', { num: true }], d.totalPallets, 'det')}</div>`;
  }
  function seccionObs(d) {
    return `<textarea class="input area-obs" id="det-observaciones" data-campo="observaciones" rows="4" maxlength="1500" placeholder="Novedades del ${T(d).uno}, si las hay" aria-label="Observaciones">${esc(d.observaciones || '')}</textarea>`;
  }
  function secciones(d) {
    return SECCIONES.map((s, i) => plegable(d, s, i,
      s.id === 'insp' ? seccionInspeccion(d) : s.id === 'carga' ? seccionCarga(d) : s.id === 'obs' ? seccionObs(d) : seccionCampos(d, s))).join('');
  }

  function vDespacho() {
    const d = buscar(vista.id);
    if (!d) return vModulo();
    const t = T(d);
    const abierto = d.estado === 'proceso';
    const pend = pendientesDe(d.id);
    const borradores = pend.filter(x => !x.listo).length;
    const marcadas = d.informeFotos || [];
    const fotoNum = name => { const m = /(\d+)/.exec(name); return m ? m[1] : ''; };
    // Botón X para eliminar (solo con el registro abierto); al tocarlo pide confirmar sobre la misma foto.
    const quitar = (id, tipo, nombre) => {
      if (!abierto || id === subiendoId) return '';
      if (st.borrar === id) return `<div class="confirm"><span>¿Eliminar foto ${esc(fotoNum(nombre))}?</span>
        <button class="si" data-act="borrarfoto" data-id="${esc(id)}" data-tipo="${tipo}" data-n="${esc(nombre)}" ${st.ocupado ? 'disabled' : ''}>Eliminar</button>
        <button class="no" data-act="noborrar">Cancelar</button></div>`;
      return `<button class="del" data-act="preborrar" data-id="${esc(id)}" aria-label="Eliminar foto ${esc(fotoNum(nombre))}">${ic('x')}</button>`;
    };
    // Estrella: la foto va al informe.
    const estrella = nombre => `<button class="star" data-act="estrella" data-n="${esc(nombre)}" aria-pressed="${marcadas.includes(nombre)}" aria-label="Usar foto ${esc(fotoNum(nombre))} en el informe">${ic('star')}</button>`;
    const tilesNube = vista.nube.slice().reverse().map(f => {
      const src = recientes.get(f.name) || f.thumb;
      return `<div class="tile ${marcadas.includes(f.name) ? 'marcada' : ''}" data-act="ver" data-id="${esc(f.id)}" role="button" tabindex="0" aria-label="Ver ${esc(f.name)} en grande">${src ? `<img src="${esc(src)}" alt="${esc(f.name)}" loading="lazy">` : `<div class="ph">${ic('image')}</div>`}<span class="num">${esc(fotoNum(f.name))}</span><span class="st ok" title="Subida">${ic('check')}</span>${estrella(f.name)}${quitar(f.id, 'nube', f.name)}</div>`;
    });
    const tilesCola = pend.slice().reverse().map(x => {
      const badge = x.error ? `<span class="st err" title="No se pudo subir">${ic('alert')}</span>`
        : !x.listo ? `<span class="st local" title="Sin subir">Sin subir</span>`
        : x.id === subiendoId ? `<span class="st up" title="Subiendo">${ic('spin')}</span>`
        : `<span class="st wait" title="Esperando">${ic('clock')}</span>`;
      return `<div class="tile ${x.listo ? 'pend' : 'draft'} ${marcadas.includes(x.archivo) ? 'marcada' : ''}" data-act="ver" data-id="${esc(x.id)}" role="button" tabindex="0" aria-label="Ver ${esc(x.archivo)} en grande"><img src="${urlsCola.get(x.id)}" alt="${esc(x.archivo)}"><span class="num">${esc(fotoNum(x.archivo))}</span>${badge}${estrella(x.archivo)}${quitar(x.id, 'cola', x.archivo)}</div>`;
    });
    const tiles = tilesCola.concat(tilesNube);
    return `<main class="wrap">
      <button class="btn ghost back" data-act="volver">${ic('back')}${st.prev === 'registros' ? 'Registros' : t.plural}</button>
      <div class="d-head">
        <span class="eyebrow">${t.Uno}</span>
        <h1>${esc(d.nombre)}</h1>
        <div class="d-meta"><span class="pill ${d.estado}">${abierto ? 'En proceso' : 'Cerrado'}</span><span>${fmtDate(d.fecha)}${d.cliente ? ' · ' + esc(d.cliente) : ''}</span></div>
        ${traza(d)}
        ${SIN_ONEDRIVE ? '' : `<div class="crumbs">${ic('folder')}OneDrive › ${esc(RAIZ)} › ${esc(t.carpeta)} › <b>${esc(d.carpeta)}</b></div>`}
      </div>
      ${abierto ? `
        <div class="row">
          <button class="btn primary big" data-act="camara" style="flex:2 1 200px">${ic('camera')}Tomar foto</button>
          <button class="btn big" data-act="galeria" style="flex:1 1 140px">${ic('image')}Desde el celular</button>
        </div>` : `
        <div class="closed-note"><span>${t.Uno} cerrado el ${fmtDate(d.cerrado)}. Si faltó alguna foto, reábrelo para agregarla.</span>
          <button class="btn" data-act="reabrir" ${st.ocupado ? 'disabled' : ''}>${ic('redo')}Reabrir para agregar fotos</button></div>`}
      ${cajaSubida(d)}
      <section class="list">
        <div class="sec-h"><h2>Fotos</h2><span class="count">${tiles.length} · <span class="cnt-star">${ic('star')} ${marcadas.length} para el informe</span></span></div>
        ${tiles.length ? `<p class="help">Toca la estrella en las fotos que van al informe.</p>` : ''}
        ${vista.error ? `<div class="note bad">${ic('alert')}<span>${esc(vista.error)} <button class="link" data-act="recargarvista">Reintentar</button></span></div>` : ''}
        ${tiles.length ? `<div class="grid">${tiles.join('')}</div>`
          : vista.cargando ? `<p class="loading">${ic('spin')}Cargando fotos…</p>`
          : `<div class="empty">Toca <b>Tomar foto</b> para empezar. Las fotos quedan en el celular hasta que tocas <b>Subir</b>.</div>`}
      </section>
      <section class="list"><span class="eyebrow">Datos del informe · se guardan solos</span>${secciones(d)}</section>
      ${sugerencias(d)}
      <div class="row">
        <button class="btn primary" data-act="informe" data-id="${esc(d.id)}" ${st.generando ? 'disabled' : ''}>${st.generando === d.id ? ic('spin') + 'Preparando informe…' : ic('doc') + 'Descargar informe (Word)'}</button>
        <button class="btn" data-act="zip" ${st.generandoZip || !tiles.length ? 'disabled' : ''}>${st.generandoZip === d.id ? ic('spin') + 'Preparando fotos…' : ic('image') + 'Descargar fotos (.zip)'}</button>
        ${SIN_ONEDRIVE ? '' : `<a class="btn" href="${esc(d.webUrl)}" target="_blank" rel="noopener">${ic('ext')}Abrir carpeta en OneDrive</a>`}
      </div>
      ${abierto ? `<button class="btn block" data-act="cerrar" ${tiles.length && !borradores && !st.ocupado ? '' : 'disabled'}>${ic('check')}Cerrar ${t.uno}</button>
        ${borradores ? `<p class="help">Para cerrar, primero sube las ${borradores} ${borradores === 1 ? 'foto' : 'fotos'} que están en el celular o elimínalas.</p>` : ''}` : ''}
      <button class="btn block ghost danger" data-act="preeliminar" ${st.ocupado ? 'disabled' : ''}>${ic('trash')}Eliminar ${t.uno}</button>
    </main>`;
  }

  // Trazabilidad: quién creó, cerró, reabrió o modificó por última vez (las horas quedan guardadas en la nube).
  function traza(d) {
    const quien = c => (rolDe(c) === 'soporte' ? 'Soporte' : rolDe(c) === 'ventas' ? 'Ventas' : c || '');
    const partes = [
      d.creadoPor && `Creado por ${quien(d.creadoPor)}`,
      d.estado === 'cerrado' && d.cerradoPor && `cerrado por ${quien(d.cerradoPor)}`,
      d.estado !== 'cerrado' && d.reabiertoPor && `reabierto por ${quien(d.reabiertoPor)}`,
      d.modificadoPor && `última modificación: ${quien(d.modificadoPor)}`
    ].filter(Boolean);
    return partes.length ? `<p class="traza">${partes.join(' · ')}</p>` : '';
  }

  // Confirmación para eliminar un despacho o traslado completo.
  function vEliminar() {
    const d = st.eliminar && buscar(st.eliminar);
    if (!d) return '';
    const t = T(d);
    const nFotos = Math.max(d.fotos || 0, vista.id === d.id ? vista.nube.length : 0) + pendientesDe(d.id).length;
    return `<div class="sheet-bg" data-act="noeliminar"><div class="sheet" data-stop>
      <span class="eyebrow">Eliminar ${t.uno}</span>
      <h2 style="font-size:20px">¿Eliminar "${esc(d.nombre)}"?</h2>
      <p class="help">Se borran el registro${nFotos ? `, sus datos y ${nFotos === 1 ? 'su foto' : `sus ${nFotos} fotos`}` : ' y sus datos'}.
        ${SIN_ONEDRIVE ? 'Esta acción no se puede deshacer.' : 'La carpeta queda en la papelera de OneDrive por si fue un error.'}</p>
      <button class="btn block primary peligro" data-act="eliminar" ${st.ocupado ? 'disabled' : ''}>${st.ocupado ? ic('spin') + 'Eliminando…' : ic('trash') + 'Sí, eliminar'}</button>
      <button class="btn block" data-act="noeliminar">Cancelar</button>
    </div></div>`;
  }

  // Aviso al cerrar si falta algo: solo informa; se puede cerrar igual.
  function vAvisoCierre() {
    const d = st.aviso && buscar(st.aviso);
    if (!d) return '';
    const falta = faltantes(d);
    const sinFotos = !(d.informeFotos || []).length;
    return `<div class="sheet-bg" data-act="cerraraviso"><div class="sheet" data-stop>
      <span class="eyebrow">Antes de cerrar</span>
      <h2 style="font-size:20px">Hay datos sin llenar</h2>
      <ul class="faltan">
        ${falta.map(({ s, a }) => `<li><button class="link" data-act="completar" data-s="${s.id}">${esc(tituloSec(s, d))}</button><span>faltan ${a[1] - a[0]}</span></li>`).join('')}
        ${sinFotos ? `<li><span>Ninguna foto marcada para el informe</span></li>` : ''}
      </ul>
      <p class="help">Puedes cerrar igual y completarlo después, o volver a llenarlo ahora.</p>
      <button class="btn block primary" data-act="cerrar" data-forzar="1">${ic('check')}Cerrar de todas formas</button>
      <button class="btn block" data-act="cerraraviso">Volver y completar</button>
    </div></div>`;
  }

  // ---------- Visor: una foto a pantalla completa, en alta calidad ----------
  const fotosGrandes = new Map(); // id → URL de la foto completa ya descargada
  function listaVisor() {
    const n = vista.nube.map(f => ({ id: f.id, nombre: f.name, tipo: 'nube', previa: recientes.get(f.name) || f.thumb }));
    const c = pendientesDe(vista.id).map(x => ({ id: x.id, nombre: x.archivo, tipo: 'cola', previa: urlsCola.get(x.id), blob: x.blob }));
    return [...n, ...c].sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { numeric: true }));
  }
  async function verFoto(i) {
    const lista = listaVisor();
    if (!lista.length) { st.visor = null; render(); return; }
    i = Math.max(0, Math.min(lista.length - 1, i));
    const f = lista[i];
    if (f.tipo === 'cola' && !fotosGrandes.has(f.id)) fotosGrandes.set(f.id, URL.createObjectURL(f.blob));
    st.visor = { i, id: f.id, cargando: !fotosGrandes.has(f.id), error: false };
    render();
    if (!st.visor.cargando) return;
    try {
      const blob = await fotoNube(f.id);
      fotosGrandes.set(f.id, URL.createObjectURL(blob));
    } catch (_) { if (st.visor && st.visor.id === f.id) st.visor.error = true; }
    if (st.visor && st.visor.id === f.id) { st.visor.cargando = false; render(); }
  }
  function vVisor() {
    if (!st.visor) return '';
    const lista = listaVisor();
    const f = lista[st.visor.i];
    if (!f) return '';
    const num = (/(\d+)/.exec(f.nombre) || [])[1] || '';
    const src = fotosGrandes.get(f.id) || f.previa;
    return `<div class="review visor" role="dialog" aria-label="Foto ${esc(num)}">
      <div class="review-top"><b>Foto ${esc(num)}</b><span class="mono">${st.visor.i + 1} de ${lista.length}</span>
        <button class="visor-x" data-act="visorcerrar" aria-label="Cerrar">${ic('x')}</button></div>
      <div class="review-img">
        ${src ? `<img src="${esc(src)}" alt="${esc(f.nombre)}">` : ''}
        ${st.visor.cargando ? `<span class="visor-carga">${ic('spin')}Cargando en alta calidad…</span>` : ''}
        ${st.visor.error ? `<span class="visor-carga">${ic('alert')}Sin señal: se muestra la versión pequeña.</span>` : ''}
      </div>
      <div class="review-bar">
        <button class="btn sec visor-nav" data-act="visorant" ${st.visor.i ? '' : 'disabled'} aria-label="Foto anterior">${ic('back')}</button>
        <button class="btn acc" data-act="visorbajar" ${st.visor.cargando ? 'disabled' : ''}>${ic('image')}Descargar foto</button>
        <button class="btn sec visor-nav" data-act="visorsig" ${st.visor.i < lista.length - 1 ? '' : 'disabled'} aria-label="Foto siguiente">${ic('chev')}</button>
      </div></div>`;
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
    app.innerHTML = encabezado(v === 'registros' || v === 'areas') + cuerpo + barra() + (st.revision ? vRevision() : '') + vVisor() + vMenu() + vAvisoCierre() + vEliminar();
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
      vista.error = e instanceof SinSenal ? 'Sin señal: no se pueden mostrar las fotos que ya están en la nube.' : 'No se pudieron cargar las fotos.';
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

  // Valor elegido en una lista desplegable propia.
  function elegirValor(id, v) {
    st.sel = null;
    if (v === OTRO) { if (id.startsWith('nuevo-')) st.nuevo[id.slice(6)] = ''; escribirOtro(id); return; }
    if (id.startsWith('nuevo-')) { st.nuevo[id.slice(6)] = v; render(); return; }
    const d = buscar(vista.id);
    if (!d) { render(); return; }
    if (id.startsWith('prod-')) {
      // Producto: se guarda y su código se llena solo.
      const i = +id.slice(5);
      const prods = (d.productos && d.productos.length ? d.productos : [{}]).map(p => ({ ...p }));
      prods[i] = { ...(prods[i] || {}), producto: v, codigo: codigoDe(d.cliente, v) || (prods[i] || {}).codigo || '' };
      guardarDetalles(d.id, { productos: prods });
    } else if (id.startsWith('det-')) {
      guardarDetalles(d.id, { [id.slice(4)]: v });
    }
    toast('Guardado');
    render();
  }

  document.addEventListener('click', async e => {
    if (e.target.closest('[data-stop]') && !e.target.closest('[data-act]')) return;
    const b = e.target.closest('[data-act]');
    // Un clic fuera de la lista abierta la cierra.
    if (st.sel && !(b && ['abrirsel', 'elegir'].includes(b.dataset.act))) { st.sel = null; render(); if (!b) return; }
    if (!b) return;
    const d = buscar(vista.id);
    switch (b.dataset.act) {
      case 'abrirsel': {
        st.sel = st.sel === b.dataset.id ? null : b.dataset.id;
        render();
        if (st.sel) setTimeout(() => document.querySelector('.sel.open .sel-pop')?.scrollIntoView({ block: 'nearest', behavior: 'smooth' }), 20);
        break;
      }
      case 'elegir': elegirValor(b.dataset.id, b.dataset.v); break;
      case 'ver': { const i = listaVisor().findIndex(f => f.id === b.dataset.id); if (i >= 0) verFoto(i); break; }
      case 'visorcerrar': st.visor = null; render(); break;
      case 'visorant': if (st.visor) verFoto(st.visor.i - 1); break;
      case 'visorsig': if (st.visor) verFoto(st.visor.i + 1); break;
      case 'visorbajar': {
        if (window.HUB_DEMO) { toast('Desde este enlace no se pueden descargar archivos. Usa la app publicada en GitHub.'); break; }
        const f = st.visor && listaVisor()[st.visor.i];
        const url = f && fotosGrandes.get(f.id);
        if (url) bajarArchivo(await (await fetch(url)).blob(), f.nombre);
        break;
      }
      case 'zip': if (d && !st.generandoZip) descargarZip(d); break;
      case 'preeliminar': st.eliminar = d.id; render(); break;
      case 'noeliminar': if (e.target === b || b.tagName === 'BUTTON') { st.eliminar = null; render(); } break;
      case 'eliminar': {
        const x = buscar(st.eliminar);
        if (!x) break;
        st.ocupado = true; render();
        try {
          try { await graph(`/me/drive/items/${x.id}`, { method: 'DELETE' }); }
          catch (err) { if (err.status !== 404) throw err; } // si la carpeta ya no estaba, igual se quita del registro
          for (const it of pendientesDe(x.id)) { await idb.quitar(it.id); urlsCola.delete(it.id); }
          await refrescarCola();
          const pend = cache.leer(clavePend()) || {};
          if (pend[x.id]) { delete pend[x.id]; cache.guardar(clavePend(), pend); }
          await cambiarRegistro(reg => { reg.despachos = reg.despachos.filter(y => y.id !== x.id); });
          st.ocupado = false; st.eliminar = null;
          ir(st.prev === 'registros' ? 'registros' : 'modulo');
          toast(`${T(x).Uno} "${x.nombre}" eliminado`);
        } catch (err) {
          st.ocupado = false; render();
          toast(errorAMensaje(err, `eliminar el ${T(x).uno}`));
        }
        break;
      }
      case 'ingresar':
        if (ELEGIR_CUENTA) { st.opcionesDemo = true; render(); } else ingresar();
        break;
      case 'demologin': iniciarDemo(b.dataset.u); break;
      case 'salir':
        if (ELEGIR_CUENTA) { if (fbEscucha) fbEscucha(); try { localStorage.removeItem(SESION_DEMO); } catch (_) {} location.reload(); }
        else if (FB) { if (fbEscucha) fbEscucha(); fbAuth.signOut().finally(() => location.reload()); }
        else salir();
        break;
      case 'migrar': if (!st.ocupado) subirDatosLocales(); break;
      case 'olvide': {
        const correo = (document.getElementById('l-correo')?.value || '').trim();
        if (!correo) { st.errorLogin = 'Escribe tu correo y vuelve a tocar "¿Olvidaste tu contraseña?".'; render(); break; }
        try {
          await fbAuth.sendPasswordResetEmail(correo);
          st.errorLogin = null; st.correo = correo; render();
          toast('Te enviamos un correo para crear una contraseña nueva.');
        } catch (_) { st.errorLogin = 'No se pudo enviar el correo. Revisa que esté bien escrito.'; st.correo = correo; render(); }
        break;
      }
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
          toast(SIN_ONEDRIVE ? `${T(st.tipo).Uno} creado` : 'Carpeta creada en OneDrive');
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
      case 'opcion': {
        const k = b.dataset.k, v = b.dataset.v;
        guardarDetalles(d.id, { [k]: d[k] === v ? '' : v });
        render();
        break;
      }
      case 'estrella': {
        const n = b.dataset.n;
        const lista = (d.informeFotos || []).slice();
        const i = lista.indexOf(n);
        if (i >= 0) lista.splice(i, 1); else lista.push(n);
        guardarDetalles(d.id, { informeFotos: lista });
        render();
        break;
      }
      case 'agregarprod': {
        const prods = (d.productos && d.productos.length ? d.productos : [{}]).map(p => ({ ...p }));
        prods.push({ producto: '', codigo: '', cantidad: '' });
        guardarDetalles(d.id, { productos: prods });
        render();
        setTimeout(() => document.querySelector(`[data-prod="producto"][data-i="${prods.length - 1}"]`)?.focus(), 30);
        break;
      }
      case 'quitarprod': {
        const prods = (d.productos || []).map(p => ({ ...p }));
        prods.splice(+b.dataset.i, 1);
        guardarDetalles(d.id, { productos: prods });
        render();
        break;
      }
      case 'cerraraviso': if (e.target === b || b.tagName === 'BUTTON') { st.aviso = null; render(); } break;
      case 'completar': {
        const s = b.dataset.s;
        st.aviso = null; st.abiertos[s] = true; render();
        setTimeout(() => document.getElementById('sec-' + s)?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50);
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
          if ((d.informeFotos || []).includes(b.dataset.n)) guardarDetalles(d.id, { informeFotos: d.informeFotos.filter(x => x !== b.dataset.n) });
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
          if ((d.informeFotos || []).includes(b.dataset.n)) guardarDetalles(d.id, { informeFotos: d.informeFotos.filter(x => x !== b.dataset.n) });
          cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); if (x) x.fotos = n; }).catch(() => {});
          toast(SIN_ONEDRIVE ? 'Foto eliminada' : 'Foto eliminada. Si fue un error, está en la papelera de OneDrive.');
        } catch (err) {
          toast(err.status === 404 ? 'Esa foto ya no estaba en la carpeta.' : errorAMensaje(err, 'eliminar la foto'));
          if (err.status === 404) vista.nube = vista.nube.filter(f => f.id !== id);
        }
        st.ocupado = false; st.borrar = null; render();
        break;
      }
      case 'cerrar': {
        if (porSubir(d.id).length) { toast('Primero sube las fotos que están en el celular.'); break; }
        // Si falta llenar algo, solo se avisa; con "Cerrar de todas formas" se cierra igual.
        if (!b.dataset.forzar && (faltantes(d).length || !(d.informeFotos || []).length)) { st.aviso = d.id; render(); break; }
        st.aviso = null;
        st.ocupado = true; render();
        try {
          const n = await contarFotos(d.id);
          await cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); x.estado = 'cerrado'; x.cerrado = new Date().toISOString(); x.cerradoPor = cuenta.username; x.fotos = Math.max(n, x.fotos || 0); });
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
          await cambiarRegistro(reg => { const x = reg.despachos.find(y => y.id === d.id); x.estado = 'proceso'; x.cerrado = null; x.reabierto = new Date().toISOString(); x.reabiertoPor = cuenta.username; });
          toast(`${T(d).Uno} reabierto`);
        } catch (err) { toast(errorAMensaje(err, `reabrir el ${T(d).uno}`)); }
        st.ocupado = false; render();
        break;
      }
      case 'alista': e.preventDefault(); delete st.libre[b.dataset.id]; render(); break;
      case 'estado': st.estado = b.dataset.e; render(); break;
      case 'limpiar': Object.assign(st, { q: '', desde: '', hasta: '', estado: 'todos' }); render(); break;
    }
  });

  // "Otro (escribir)" en una lista: se cambia por un campo de texto para escribir el valor.
  function escribirOtro(id) {
    st.libre[id] = true;
    render();
    setTimeout(() => document.getElementById(id)?.focus(), 30);
  }

  document.addEventListener('input', e => {
    const t = e.target;
    if (t.value === OTRO && t.id) { escribirOtro(t.id); if (t.id.startsWith('nuevo-')) st.nuevo[t.dataset.campo] = ''; return; }
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
    if (t.value === OTRO) return; // lo atiende el evento "input": cambia la lista por un campo para escribir
    if (t.id && t.id.startsWith('det-')) {
      const v = t.value.trim();
      if ((d[t.dataset.campo] || '') !== v) {
        guardarDetalles(d.id, { [t.dataset.campo]: v });
        toast('Guardado');
        if (t.dataset.campo === 'cliente') render(); // cambia la lista de productos
      }
    } else if (t.dataset.prod) {
      // Fila de producto: al elegir el producto se llena solo su código.
      const i = +t.dataset.i;
      const prods = (d.productos && d.productos.length ? d.productos : [{}]).map(p => ({ ...p }));
      prods[i] = { ...(prods[i] || {}), [t.dataset.prod]: t.value.trim() };
      if (t.dataset.prod === 'producto') {
        const cod = codigoDe(d.cliente, t.value);
        if (cod) {
          prods[i].codigo = cod;
          const c = document.querySelector(`[data-prod="codigo"][data-i="${i}"]`);
          if (c) c.value = cod;
        }
      }
      guardarDetalles(d.id, { productos: prods });
      toast('Guardado');
    } else if (t.dataset.obs) {
      guardarDetalles(d.id, { inspeccion: { [t.dataset.obs]: { obs: t.value.trim() } } });
      toast('Guardado');
    }
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.target.id && e.target.id.startsWith('nuevo-')) document.getElementById('crear').click();
    if (e.key === 'Enter' && (e.target.id?.startsWith('det-') || e.target.dataset?.obs)) e.target.blur();
    if (e.key === 'Escape' && (st.menu || st.sel || st.eliminar || st.visor)) { st.menu = false; st.sel = null; st.eliminar = null; st.visor = null; render(); }
    if (st.visor && e.key === 'ArrowLeft') verFoto(st.visor.i - 1);
    if (st.visor && e.key === 'ArrowRight') verFoto(st.visor.i + 1);
    if (e.key === 'Enter' && e.target.classList?.contains('tile')) e.target.click();
  });

  // ---------- Arranque ----------
  async function iniciarDemo(clave) {
    cuenta = DEMO_CUENTAS[clave];
    if (!cuenta) { ir('ingreso'); return; }
    if (clave === 'otro') { ir('sinacceso'); return; }
    try { localStorage.setItem(SESION_DEMO, clave); } catch (_) {}
    st.rol = clave;
    if (FB) { // Firebase sin cuentas: se elige Ventas o Soporte, pero los datos van a la nube
      registro = aplicarPendientes(cache.leer('registro.' + usuarioDatos()) || { despachos: [] });
      await refrescarCola();
      ir('areas');
      fbEscuchar();
      procesarCola();
      revisarDatosLocales();
      return;
    }
    await demoCargar();
    await cargarRegistro();
    await refrescarCola();
    ir('areas');
    procesarCola();
  }

  const rolDe = correo => {
    const permitidos = {};
    Object.keys(CFG.correosPermitidos || {}).forEach(k => { permitidos[k.toLowerCase()] = CFG.correosPermitidos[k]; });
    return permitidos[(correo || '').toLowerCase()] || null;
  };
  async function entrarFirebase(u) {
    st.rol = rolDe(u.email);
    cuenta = { username: (u.email || '').toLowerCase(), name: st.rol === 'soporte' ? 'Soporte' : st.rol === 'ventas' ? 'Ventas' : u.email };
    if (!st.rol) { ir('sinacceso'); return; }
    registro = aplicarPendientes(cache.leer('registro.' + usuarioDatos()) || { despachos: [] });
    await refrescarCola();
    ir('areas');
    fbEscuchar();
    procesarCola();
    revisarDatosLocales();
  }
  // Registros hechos antes de la nube (modo prueba, solo en este dispositivo): se ofrece subirlos.
  function revisarDatosLocales() {
    const d = cache.leer('demo.datos.v6');
    if (!d || cache.leer('migrado.v6')) return;
    const n = (JSON.parse(d.registro || '{}').despachos || []).length;
    if (n) { st.migrar = n; render(); } else cache.guardar('migrado.v6', true);
  }
  async function subirDatosLocales() {
    const d = cache.leer('demo.datos.v6');
    const regs = JSON.parse(d.registro || '{}').despachos || [];
    st.ocupado = true; render();
    try {
      const lote = fbDb.batch();
      regs.forEach(x => lote.set(fbDb.collection('registros').doc(x.id), limpio({ ...x, creadoPor: x.creadoPor || cuenta.username, creado: x.creado || x.fecha })));
      await conPlazo(lote.commit(), 60000);
      let fotos = 0;
      for (const c of d.carpetas || []) {
        if (!regs.some(x => x.id === c.id)) continue;
        for (const f of c.fotos || []) {
          const blob = await idb.leer('demofoto.' + f.id).catch(() => null);
          if (!blob) continue;
          await fbSubirFoto({ id: f.id, despachoId: c.id, archivo: f.name, blob });
          fotos++;
        }
      }
      cache.guardar('migrado.v6', true);
      st.migrar = 0;
      toast(`Listo: ${regs.length} ${regs.length === 1 ? 'registro' : 'registros'} y ${fotos} fotos ya están en la nube.`);
    } catch (e) {
      console.error(e);
      toast(errorAMensaje(e, 'subir los registros de este dispositivo'));
    }
    st.ocupado = false; render();
  }
  document.addEventListener('submit', async e => {
    if (e.target.id !== 'f-login') return;
    e.preventDefault();
    const correo = document.getElementById('l-correo').value.trim();
    const clave = document.getElementById('l-clave').value;
    st.correo = correo;
    if (!correo || !clave) { st.errorLogin = 'Escribe tu correo y tu contraseña.'; render(); return; }
    st.ocupado = true; st.errorLogin = null; render();
    try {
      const r = await fbAuth.signInWithEmailAndPassword(correo, clave);
      st.ocupado = false;
      await entrarFirebase(r.user);
    } catch (err) {
      st.ocupado = false;
      const c = err.code || '';
      st.errorLogin = /invalid-credential|wrong-password|user-not-found|invalid-email/.test(c) ? 'Correo o contraseña incorrectos.'
        : /too-many-requests/.test(c) ? 'Demasiados intentos. Espera unos minutos y vuelve a probar.'
        : /network/.test(c) ? 'No hay conexión a internet. Para ingresar la primera vez se necesita señal.'
        : 'No se pudo iniciar sesión. Intenta de nuevo.';
      render();
    }
  });

  async function arrancar() {
    if (FB) {
      if (!CFG.firebase || !CFG.firebase.apiKey || /PEGAR/.test(CFG.firebase.apiKey)) { ir('falla', { error: 'Falta configurar la conexión a la nube (config.js). Avisa a soporte.' }); return; }
      let u;
      try { u = await iniciarFirebase(); }
      catch (e) { console.error(e); ir('falla', { error: 'No se pudo conectar con la nube. Revisa la conexión a internet e intenta de nuevo.' }); return; }
      if (FB_CUENTAS) { if (u) await entrarFirebase(u); else ir('ingreso'); return; }
      let guardada = null;
      try { guardada = localStorage.getItem(SESION_DEMO); } catch (_) {}
      if (guardada && DEMO_CUENTAS[guardada]) await iniciarDemo(guardada); else ir('ingreso');
      return;
    }
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

    st.rol = rolDe(cuenta.username);
    if (!st.rol) { ir('sinacceso'); return; }

    registro = aplicarPendientes(cache.leer('registro.' + usuarioDatos()) || { despachos: [] });
    await refrescarCola();
    ir('areas');
    try { await cargarRegistro(); refrescar(); }
    catch (e) { if (!(e instanceof SinSenal)) { console.error(e); toast('No se pudo actualizar el registro desde OneDrive.'); } }
    procesarCola();
  }

  if (!window.HUB_DEMO && 'serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  arrancar();
})();
