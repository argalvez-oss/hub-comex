/* Hub Comex — genera el informe de despacho / traslado en Word (.docx).
 *
 * Formato final aprobado por Ventas (oct. 2026). Todo lo que se llenó en la app sale en el informe;
 * lo que quedó vacío se puede completar en Word. Las fotos marcadas con estrella en la app se colocan solas.
 * Para cambiar el formato del informe, se edita solo este archivo.
 *
 * Uso: HCInforme.crear(docx, registro, opciones) → docx.Document   (docx = librería "docx" 9.x)
 *   opciones: { codigo, revision, empresa, logo: { data, w, h }, fotos: [{ data, w, h, nombre }] }
 */
(function (raiz) {
  'use strict';

  const ANCHO = 10368; // ancho útil de la hoja carta con márgenes de 0,65" (en DXA)
  const COLOR = { marca: '17365D', seccion: '1F4E78', etiqueta: 'E3EEF8', suave: 'F5F7FA', borde: 'B7C3D0', gris: '7F8C99', texto: '1F2937' };
  const INSPECCION = [
    ['exterior', 'Exterior, puertas y estructura'],
    ['interior', 'Interior limpio, seco y sin olores'],
    ['piso', 'Piso, paredes y techo íntegros'],
    ['refrigeracion', 'Refrigeración operativa'],
    ['llantas', 'Llantas, luces y elementos de seguridad'],
    ['apta', 'Unidad apta para carga de alimentos']
  ];

  const pad = n => String(n).padStart(2, '0');
  const fecha = v => { if (!v) return ''; const d = new Date(v); return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; };
  const hora = v => { if (!v) return ''; const d = new Date(v); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
  // Ajusta una imagen dentro de un recuadro (en píxeles) sin deformarla.
  const encajar = (w, h, maxW, maxH) => { const k = Math.min(maxW / w, maxH / h); return { width: Math.round(w * k), height: Math.round(h * k) }; };

  function crear(D, reg, op) {
    op = op || {};
    const { Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType,
      VerticalAlign, HeightRule, Header, Footer, PageNumber, ExternalHyperlink, TabStopType } = D;
    const esTraslado = reg.tipo === 'traslado';
    const titulo = esTraslado ? 'INFORME DE TRASLADO' : 'INFORME DE DESPACHO';
    const v = k => (reg[k] == null ? '' : String(reg[k]));

    // ---- piezas básicas ----
    const linea = { style: BorderStyle.SINGLE, size: 4, color: COLOR.borde };
    const bordes = { top: linea, bottom: linea, left: linea, right: linea, insideHorizontal: linea, insideVertical: linea };
    const texto = (t, o = {}) => new TextRun({ text: t == null ? '' : String(t), size: o.size || 18, bold: o.bold, color: o.color || COLOR.texto, italics: o.italics, font: o.font });
    const parrafo = (hijos, o = {}) => new Paragraph({ children: Array.isArray(hijos) ? hijos : [hijos], alignment: o.align, spacing: { before: o.antes || 0, after: o.despues || 0 }, keepNext: o.keepNext });
    // o.keep: mantiene la fila junto con la siguiente (para que un bloque no se parta entre páginas)
    const celda = (contenido, ancho, o = {}) => new TableCell({
      children: Array.isArray(contenido) && contenido[0] instanceof Paragraph ? contenido : [parrafo(contenido, { align: o.align, keepNext: o.keep })],
      width: { size: ancho, type: WidthType.DXA },
      columnSpan: o.span,
      shading: o.fondo ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fondo } : undefined,
      verticalAlign: o.vAlign || VerticalAlign.CENTER,
      margins: { top: 70, bottom: 70, left: 110, right: 110 }
    });
    const tabla = (anchos, filas) => new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: anchos, rows: filas, borders: bordes });
    const fila = (celdas, o = {}) => new TableRow({ children: celdas, cantSplit: true, height: o.alto ? { value: o.alto, rule: o.exacto ? HeightRule.EXACT : HeightRule.ATLEAST } : undefined, tableHeader: o.encabezado });
    const espacio = (n = 120) => new Paragraph({ spacing: { before: 0, after: n }, children: [] });
    const casilla = marcada => new TextRun({ text: marcada ? '☒' : '☐', size: 20, font: 'Segoe UI Symbol', color: COLOR.marca });
    const etiqueta = t => texto(t, { bold: true, size: 17, color: COLOR.marca });
    const valor = x => x instanceof TextRun || Array.isArray(x) ? x : texto(x);

    let numero = 0;
    const seccion = nombre => tabla([ANCHO], [fila([celda(texto(`${++numero}. ${nombre}`, { bold: true, color: 'FFFFFF', size: 19 }), ANCHO, { fondo: COLOR.seccion, keep: true })])]);
    // Tabla de pares etiqueta / valor en 4 columnas. Una fila con solo 2 elementos ocupa todo el ancho.
    const A4 = [1950, 3234, 1950, 3234];
    const pares = filas => tabla(A4, filas.map((f, i) => {
      const keep = i < filas.length - 1;
      const celdas = f.length === 2
        ? [celda(etiqueta(f[0]), A4[0], { fondo: COLOR.etiqueta, keep }), celda(valor(f[1]), A4[1] + A4[2] + A4[3], { span: 3, keep })]
        : [celda(etiqueta(f[0]), A4[0], { fondo: COLOR.etiqueta, keep }), celda(valor(f[1]), A4[1], { keep }),
           celda(etiqueta(f[2]), A4[2], { fondo: COLOR.etiqueta, keep }), celda(valor(f[3]), A4[3], { keep })];
      return fila(celdas, { alto: 370 });
    }));

    // ---- encabezado (con logo) y pie ----
    const H = [2200, 5368, 2800];
    const logo = op.logo && op.logo.data
      ? new ImageRun({ type: 'png', data: op.logo.data, transformation: encajar(op.logo.w, op.logo.h, 90, 72) })
      : texto('HEALTHY FOOD', { bold: true, color: '7E9A3F', size: 24 });
    const encabezado = new Header({
      children: [
        tabla(H, [
          fila([
            celda([parrafo(logo, { align: AlignmentType.CENTER })], H[0]),
            celda([parrafo(texto(titulo, { bold: true, color: COLOR.marca, size: 28 }), { align: AlignmentType.CENTER })], H[1]),
            celda([
              parrafo([texto('Código: ', { bold: true, size: 15 }), texto(op.codigo || 'HF-LOG-___', { size: 15 })]),
              parrafo([texto('Revisión: ', { bold: true, size: 15 }), texto(op.revision || '___', { size: 15 })], { antes: 40 }),
              parrafo([texto('Página: ', { bold: true, size: 15 }), new TextRun({ children: [PageNumber.CURRENT], size: 15 }), texto(' de ', { size: 15 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15 })], { antes: 40 })
            ], H[2])
          ], { alto: 1100 })
        ]),
        espacio(0)
      ]
    });
    const pie = new Footer({
      children: [new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: COLOR.borde, space: 4 } },
        tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }],
        children: [
          texto(`${op.empresa || 'Healthy Food Marcalman Ecuador S.A.S.'} | ${esTraslado ? 'Informe de traslado' : 'Informe de despacho'}`, { size: 14, color: COLOR.gris }),
          texto(`\tGenerado con Hub Comex el ${fecha(new Date())}`, { size: 14, color: COLOR.gris })
        ]
      })]
    });

    const linkCarpeta = (size = 18) => reg.webUrl && /^https?:/.test(reg.webUrl)
      ? new ExternalHyperlink({ link: reg.webUrl, children: [new TextRun({ text: reg.carpeta || reg.nombre, style: 'Hyperlink', size })] })
      : texto(reg.carpeta || reg.nombre, { size });

    // ---- 1. Datos generales ----
    const datos = [
      seccion('DATOS GENERALES'),
      pares([
        ['Operación', [casilla(!esTraslado), texto(' Exportación     '), casilla(esTraslado), texto(' Traslado')], 'Fecha', fecha(reg.fecha)],
        ['Cliente', v('cliente'), 'Hora de inicio', hora(reg.fecha)],
        ['Destino', v('destino'), 'Hora de fin', hora(reg.cerrado)],
        ['Lugar de carga', v('lugarCarga'), 'Destino final', v('destinoFinal')],
        ['Pedido / OC', v('pedido'), 'Factura / DAE', v('factura')],
        ['Responsable', v('responsable'), 'C.I.', v('ci')],
        ['Carpeta de fotos', [linkCarpeta()]]
      ]),
      espacio()
    ];

    // ---- 2. Transporte y contenedor ----
    const temp = k => (v(k) ? `${v(k)}` : '');
    const transporte = [
      seccion('TRANSPORTE Y CONTENEDOR'),
      pares([
        ['Transportista', v('transportista'), 'Conductor', v('conductor')],
        ['Cédula', v('cedula'), 'Placa', v('placa')],
        ['Contenedor', v('contenedor'), 'Tipo', [casilla(reg.tipoUnidad === 'reefer'), texto(' Reefer     '), casilla(reg.tipoUnidad === 'camion'), texto(' Camión')]],
        ['Naviera', v('naviera'), 'Capacidad', v('capacidad')],
        ['T. inicial (°C)', temp('tInicial'), 'T. final (°C)', temp('tFinal')],
        ['Set point (°C)', temp('setPoint'), 'Termógrafo', v('termografo')]
      ]),
      espacio()
    ];

    // ---- 3. Inspección previa ----
    const I = [4300, 800, 800, ANCHO - 5900];
    const insp = reg.inspeccion || {};
    const inspeccion = [
      seccion('INSPECCIÓN PREVIA DE LA UNIDAD'),
      tabla(I, [
        fila([
          celda(etiqueta('Elemento'), I[0], { fondo: COLOR.etiqueta, keep: true }),
          celda(etiqueta('Sí'), I[1], { fondo: COLOR.etiqueta, align: AlignmentType.CENTER, keep: true }),
          celda(etiqueta('No'), I[2], { fondo: COLOR.etiqueta, align: AlignmentType.CENTER, keep: true }),
          celda(etiqueta('Observación'), I[3], { fondo: COLOR.etiqueta, keep: true })
        ], { encabezado: true }),
        ...INSPECCION.map(([clave, nombre], i) => {
          const r = insp[clave] || {};
          const keep = i < INSPECCION.length - 1;
          return fila([
            celda(texto(nombre), I[0], { keep }),
            celda(casilla(r.v === 'si'), I[1], { align: AlignmentType.CENTER, keep }),
            celda(casilla(r.v === 'no'), I[2], { align: AlignmentType.CENTER, keep }),
            celda(texto(r.obs || ''), I[3], { keep })
          ], { alto: 360 });
        })
      ]),
      espacio()
    ];

    // ---- 4. Estiba, sellos y cierre ----
    const sellos = [
      seccion('ESTIBA, SELLOS Y CIERRE'),
      pares([
        ['Sello naviera', v('selloNaviera'), 'Sello botella', v('selloBotella')],
        ['Sello cable', v('selloCable'), 'Sticker', v('sticker')],
        ['Otros sellos', v('otrosSellos')]
      ]),
      espacio()
    ];

    // ---- 5. Detalle de carga ----
    const C = [700, 5000, 2000, ANCHO - 7700];
    const productos = (reg.productos || []).filter(p => p && (p.producto || p.codigo || p.cantidad));
    const filasCarga = Math.max(7, productos.length);
    const carga = [
      seccion('DETALLE DE CARGA'),
      tabla(C, [
        fila(['N.º', 'Producto', 'Código', 'Cantidad'].map((t, i) => celda(etiqueta(t), C[i], { fondo: COLOR.etiqueta, align: AlignmentType.CENTER, keep: true })), { encabezado: true }),
        ...Array.from({ length: filasCarga }, (_, n) => {
          const p = productos[n] || {};
          return fila([
            celda(texto(n + 1, { color: COLOR.gris }), C[0], { align: AlignmentType.CENTER, fondo: n % 2 ? COLOR.suave : undefined }),
            celda(texto(p.producto || ''), C[1], { fondo: n % 2 ? COLOR.suave : undefined }),
            celda(texto(p.codigo || ''), C[2], { align: AlignmentType.CENTER, fondo: n % 2 ? COLOR.suave : undefined }),
            celda(texto(p.cantidad || ''), C[3], { align: AlignmentType.CENTER, fondo: n % 2 ? COLOR.suave : undefined })
          ], { alto: 360 });
        })
      ]),
      pares([['Total cajas', v('totalCajas'), 'Total pallets', v('totalPallets')]]),
      espacio()
    ];

    // ---- 6. Observaciones ----
    const lineasObs = v('observaciones').split(/\r?\n/);
    const observaciones = [
      seccion('OBSERVACIONES'),
      tabla([ANCHO], [fila([celda(lineasObs.map(l => parrafo(texto(l))), ANCHO, { vAlign: VerticalAlign.TOP })], { alto: 1300 })]),
      espacio()
    ];

    // ---- 7. Registro fotográfico: fotos marcadas en la app, 2 por fila; si no hay, espacios para pegarlas ----
    const F = [ANCHO / 2, ANCHO / 2];
    const fotosApp = op.fotos || [];
    const totalEspacios = fotosApp.length ? fotosApp.length + (fotosApp.length % 2) : 10;
    const filasFotos = [];
    for (let i = 0; i < totalEspacios; i += 2) {
      filasFotos.push(fila(F.map((w, k) => {
        const f = fotosApp[i + k];
        const contenido = f
          ? parrafo(new ImageRun({ type: 'jpg', data: f.data, transformation: encajar(f.w, f.h, 330, 220) }), { align: AlignmentType.CENTER, keepNext: true })
          : parrafo(texto(fotosApp.length ? '' : 'Insertar fotografía', { color: 'A6B1BC', italics: true }), { align: AlignmentType.CENTER, keepNext: true });
        return celda([contenido], w, { fondo: f ? undefined : COLOR.suave });
      }), { alto: 3350, exacto: true }));
      filasFotos.push(fila(F.map((w, k) => celda([parrafo([etiqueta(i + k < totalEspacios && (fotosApp[i + k] || !fotosApp.length) ? `Foto ${i + k + 1}: ` : ''), texto('')])], w)), { alto: 380 }));
    }
    const fotos = [
      new Paragraph({ pageBreakBefore: true, children: [] }),
      seccion('REGISTRO FOTOGRÁFICO'),
      new Paragraph({
        spacing: { before: 80, after: 100 },
        children: [texto('Todas las fotografías están en la carpeta ', { size: 16, color: COLOR.gris }), linkCarpeta(16),
          texto(fotosApp.length ? '.' : '. Para agregar más espacios, copia y pega una fila de la tabla.', { size: 16, color: COLOR.gris })]
      }),
      tabla(F, filasFotos),
      espacio()
    ];

    // ---- 8. Personal a cargo (al final; Bodega lleva 3 nombres) ----
    const S = [3456, 3456, 3456];
    // Espacio para firmar (si se requiere), la línea y el nombre debajo.
    const nombre = (n, alto) => [
      parrafo(texto(''), { antes: alto }),
      parrafo(texto('________________________________', { color: COLOR.borde }), { align: AlignmentType.CENTER }),
      parrafo([etiqueta('Nombre: '), texto(n || '')], { align: AlignmentType.CENTER, antes: 20 })
    ];
    const personal = [
      new Paragraph({ keepNext: true, spacing: { before: 0, after: 0 }, children: [] }),
      seccion(esTraslado ? 'PERSONAL A CARGO DEL TRASLADO' : 'PERSONAL A CARGO DEL DESPACHO'),
      tabla(S, [
        fila(['Jefe de Calidad', 'Responsable de sellos', 'Bodega'].map((r, i) => celda(texto(r, { bold: true, color: COLOR.marca }), S[i], { fondo: COLOR.etiqueta, align: AlignmentType.CENTER, keep: true })), { alto: 380 }),
        fila([
          celda([...nombre(v('jefeCalidad'), 900), parrafo(texto(''), { despues: 60 })], S[0], { vAlign: VerticalAlign.BOTTOM }),
          celda([...nombre(v('respSellos'), 900), parrafo(texto(''), { despues: 60 })], S[1], { vAlign: VerticalAlign.BOTTOM }),
          celda([...nombre(v('bodega1'), 300), ...nombre(v('bodega2'), 300), ...nombre(v('bodega3'), 300), parrafo(texto(''), { despues: 60 })], S[2], { vAlign: VerticalAlign.BOTTOM })
        ], { alto: 1900 })
      ])
    ];

    return new Document({
      creator: 'Hub Comex',
      title: `${esTraslado ? 'Informe de traslado' : 'Informe de despacho'} - ${reg.nombre}`,
      styles: { default: { document: { run: { font: 'Aptos', size: 18, color: COLOR.texto } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1900, right: 936, bottom: 1000, left: 936, header: 500, footer: 450 } } },
        headers: { default: encabezado },
        footers: { default: pie },
        children: [...datos, ...transporte, ...inspeccion, ...sellos, ...carga, ...observaciones, ...fotos, ...personal]
      }]
    });
  }

  const nombreArchivo = reg => `${reg.tipo === 'traslado' ? 'Informe de traslado' : 'Informe de despacho'} - ${reg.carpeta || reg.nombre}.docx`
    .replace(/[\\/:*?"<>|]/g, '-');

  const api = { crear, nombreArchivo, INSPECCION };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  raiz.HCInforme = api;
})(typeof window !== 'undefined' ? window : globalThis);
