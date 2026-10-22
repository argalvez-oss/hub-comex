/* Hub Comex — genera el informe de despacho / traslado en Word (.docx).
 *
 * Lo que se llenó en la app (datos generales e inspección previa) sale completo; transporte, carga,
 * sellos, observaciones, fotografías y firmas quedan como campos vacíos para completar en Word.
 * Para cambiar el formato del informe, se edita solo este archivo.
 *
 * Uso: HCInforme.crear(docx, registro, opciones) → docx.Document   (docx = librería "docx" 9.x)
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

  function crear(D, reg, op) {
    op = op || {};
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType,
      VerticalAlign, HeightRule, Header, Footer, PageNumber, ExternalHyperlink, TabStopType } = D;
    const esTraslado = reg.tipo === 'traslado';
    const titulo = esTraslado ? 'INFORME DE TRASLADO' : 'INFORME DE DESPACHO';

    // ---- piezas básicas ----
    const linea = { style: BorderStyle.SINGLE, size: 4, color: COLOR.borde };
    const bordes = { top: linea, bottom: linea, left: linea, right: linea, insideHorizontal: linea, insideVertical: linea };
    const sinBordes = (() => { const n = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }; return { top: n, bottom: n, left: n, right: n, insideHorizontal: n, insideVertical: n }; })();
    const texto = (t, o = {}) => new TextRun({ text: t == null ? '' : String(t), size: o.size || 18, bold: o.bold, color: o.color || COLOR.texto, italics: o.italics, font: o.font });
    const parrafo = (hijos, o = {}) => new Paragraph({ children: Array.isArray(hijos) ? hijos : [hijos], alignment: o.align, spacing: { before: o.antes || 0, after: o.despues || 0 }, keepNext: o.keepNext });
    // o.keep: mantiene la fila junto con la siguiente (para que una tabla no se parta entre páginas)
    const celda = (contenido, ancho, o = {}) => new TableCell({
      children: Array.isArray(contenido) && contenido[0] instanceof Paragraph ? contenido : [parrafo(contenido, { align: o.align, keepNext: o.keep })],
      width: { size: ancho, type: WidthType.DXA },
      columnSpan: o.span,
      shading: o.fondo ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fondo } : undefined,
      verticalAlign: o.vAlign || VerticalAlign.CENTER,
      margins: { top: 70, bottom: 70, left: 110, right: 110 }
    });
    const tabla = (anchos, filas, o = {}) => new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: anchos, rows: filas, borders: o.bordes || bordes });
    const fila = (celdas, o = {}) => new TableRow({ children: celdas, cantSplit: true, height: o.alto ? { value: o.alto, rule: o.exacto ? HeightRule.EXACT : HeightRule.ATLEAST } : undefined, tableHeader: o.encabezado });
    const espacio = (n = 120) => new Paragraph({ spacing: { before: 0, after: n }, children: [] });
    const casilla = marcada => new TextRun({ text: marcada ? '☒' : '☐', size: 20, font: 'Segoe UI Symbol', color: COLOR.marca });

    let numero = 0;
    const seccion = nombre => tabla([ANCHO], [fila([celda(texto(`${++numero}. ${nombre}`, { bold: true, color: 'FFFFFF', size: 19 }), ANCHO, { fondo: COLOR.seccion, keep: true })])]);
    const etiqueta = t => texto(t, { bold: true, size: 17, color: COLOR.marca });
    // Tabla de pares etiqueta / valor en 4 columnas
    const A4 = [1950, 3234, 1950, 3234];
    const valor = v => v instanceof TextRun || Array.isArray(v) ? v : texto(v);
    const pares = filas => tabla(A4, filas.map(([e1, v1, e2, v2], i) => {
      const keep = i < filas.length - 1;
      return fila([
        celda(etiqueta(e1), A4[0], { fondo: COLOR.etiqueta, keep }), celda(valor(v1), A4[1], { keep }),
        celda(etiqueta(e2), A4[2], { fondo: COLOR.etiqueta, keep }), celda(valor(v2), A4[3], { keep })
      ], { alto: 370 });
    }));

    // ---- encabezado y pie ----
    const H = [2600, 4968, 2800];
    const encabezado = new Header({
      children: [
        tabla(H, [
          fila([
            celda([parrafo(texto('HEALTHY FOOD', { bold: true, color: 'FFFFFF', size: 26 }), { align: AlignmentType.CENTER })], H[0], { fondo: COLOR.marca }),
            celda([
              parrafo(texto(titulo, { bold: true, color: COLOR.marca, size: 26 }), { align: AlignmentType.CENTER }),
              parrafo(texto(esTraslado ? 'Registro de traslado de producto' : 'Registro de despacho de producto', { size: 16, color: COLOR.gris }), { align: AlignmentType.CENTER })
            ], H[1]),
            celda([
              parrafo([texto('Código: ', { bold: true, size: 15 }), texto(op.codigo || 'HF-LOG-___', { size: 15 })]),
              parrafo([texto('Revisión: ', { bold: true, size: 15 }), texto(op.revision || '___', { size: 15 })], { antes: 40 }),
              parrafo([texto('Página: ', { bold: true, size: 15 }), new TextRun({ children: [PageNumber.CURRENT], size: 15 }), texto(' de ', { size: 15 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15 })], { antes: 40 })
            ], H[2])
          ], { alto: 820 })
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

    // ---- 1. Datos generales ----
    const enlace = reg.webUrl && /^https?:/.test(reg.webUrl)
      ? new ExternalHyperlink({ link: reg.webUrl, children: [new TextRun({ text: reg.carpeta || reg.nombre, style: 'Hyperlink', size: 18 })] })
      : texto(reg.carpeta || reg.nombre);
    const datos = [
      seccion('DATOS GENERALES'),
      pares([
        ['Operación', [casilla(!esTraslado), texto(' Exportación     '), casilla(esTraslado), texto(' Traslado')], 'Fecha', fecha(reg.fecha)],
        ['Cliente', reg.cliente, 'Hora de inicio', hora(reg.fecha)],
        ['Destino', reg.destino, 'Hora de fin', hora(reg.cerrado)],
        ['Lugar de carga', reg.lugarCarga, 'Destino final', ''],
        ['Pedido / OC', reg.pedido, 'Factura / DAE', ''],
        ['Responsable', '', 'C.I.', '']
      ]),
      tabla([A4[0], ANCHO - A4[0]], [fila([
        celda(etiqueta('Carpeta de fotos'), A4[0], { fondo: COLOR.etiqueta }),
        celda([new Paragraph({ children: [enlace] })], ANCHO - A4[0])
      ], { alto: 400 })]),
      espacio()
    ];

    // ---- 2. Transporte y contenedor (se llena en Word) ----
    const transporte = [
      seccion('TRANSPORTE Y CONTENEDOR'),
      pares([
        ['Transportista', '', 'Conductor', ''],
        ['Cédula', '', 'Placa', ''],
        ['Contenedor', '', 'Tipo', [casilla(false), texto(' Reefer     '), casilla(false), texto(' Camión')]],
        ['Naviera', '', 'Capacidad', ''],
        ['Hora de ingreso', '', 'Fin de carga', ''],
        ['T. inicial (°C)', '', 'T. final (°C)', ''],
        ['Set point (°C)', '', 'Termógrafo', '']
      ]),
      espacio()
    ];

    // ---- 3. Inspección previa (desde la app) ----
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

    // ---- 4. Detalle de la carga (se llena en Word) ----
    const C = [600, 2468, 1300, 1500, 1000, 1000, 1200, 1300];
    const cab = ['N.º', 'Producto', 'Lote', 'Presentación', 'Cajas', 'Pallets', 'Peso (kg)', 'Temp. (°C)'];
    const T = [2100, 1356, 2100, 1356, 2100, 1356];
    const carga = [
      seccion('DETALLE DE LA CARGA'),
      tabla(C, [
        fila(cab.map((t, i) => celda(etiqueta(t), C[i], { fondo: COLOR.etiqueta, align: AlignmentType.CENTER })), { encabezado: true }),
        ...Array.from({ length: 7 }, (_, n) => fila(C.map((w, i) => celda(i === 0 ? texto(n + 1, { color: COLOR.gris }) : texto(''), w, { align: AlignmentType.CENTER, fondo: n % 2 ? COLOR.suave : undefined })), { alto: 360 }))
      ]),
      tabla(T, [fila([
        celda(etiqueta('Total cajas'), T[0], { fondo: COLOR.etiqueta }), celda(texto(''), T[1]),
        celda(etiqueta('Total pallets'), T[2], { fondo: COLOR.etiqueta }), celda(texto(''), T[3]),
        celda(etiqueta('Peso total (kg)'), T[4], { fondo: COLOR.etiqueta }), celda(texto(''), T[5])
      ], { alto: 380 })]),
      espacio()
    ];

    // ---- 5. Estiba, sellos y cierre (se llena en Word) ----
    const sellos = [
      seccion('ESTIBA, SELLOS Y CIERRE'),
      pares([
        ['Sello naviera', '', 'Sello botella', ''],
        ['Sello cable', '', 'Sticker', ''],
        ['Otros sellos', '', 'Documentos adjuntos', '']
      ]),
      espacio()
    ];

    // ---- 6. Observaciones ----
    const observaciones = [
      seccion('OBSERVACIONES'),
      tabla([ANCHO], [fila([celda(texto(''), ANCHO, { vAlign: VerticalAlign.TOP })], { alto: 1300 })])
    ];

    // ---- 7. Registro fotográfico (ventas pega las fotos en el orden que quiera) ----
    const F = [ANCHO / 2, ANCHO / 2];
    const filasFotos = [];
    for (let i = 0; i < 5; i++) {
      filasFotos.push(fila(F.map(w => celda([parrafo(texto('Insertar fotografía', { color: 'A6B1BC', italics: true }), { align: AlignmentType.CENTER })], w, { fondo: COLOR.suave })), { alto: 3350, exacto: true }));
      filasFotos.push(fila(F.map((w, k) => celda([parrafo([etiqueta(`Foto ${i * 2 + k + 1}: `), texto('')])], w)), { alto: 380 }));
    }
    const fotos = [
      new Paragraph({ pageBreakBefore: true, children: [] }),
      seccion('REGISTRO FOTOGRÁFICO'),
      new Paragraph({
        spacing: { before: 80, after: 100 },
        children: [texto('Fotografías del registro disponibles en la carpeta ', { size: 16, color: COLOR.gris }),
          reg.webUrl && /^https?:/.test(reg.webUrl)
            ? new ExternalHyperlink({ link: reg.webUrl, children: [new TextRun({ text: reg.carpeta || reg.nombre, style: 'Hyperlink', size: 16 })] })
            : texto(reg.carpeta || reg.nombre, { size: 16, bold: true, color: COLOR.gris }),
          texto('. Para agregar más espacios, copia y pega una fila de la tabla.', { size: 16, color: COLOR.gris })]
      }),
      tabla(F, filasFotos),
      espacio()
    ];

    // ---- 8. Firmas (siempre al final) ----
    const S = [3456, 3456, 3456];
    const roles = ['Jefe de Calidad', 'Responsable de sellos', 'Bodega'];
    const lineaFirma = t => parrafo([etiqueta(t), texto(' ______________________________', { color: COLOR.borde })], { antes: 120 });
    const firmas = [
      new Paragraph({ keepNext: true, spacing: { before: 0, after: 0 }, children: [] }),
      seccion('FIRMAS DE RESPONSABILIDAD'),
      tabla(S, [
        fila(roles.map((r, i) => celda(texto(r, { bold: true, color: COLOR.marca }), S[i], { fondo: COLOR.etiqueta, align: AlignmentType.CENTER, keep: true })), { alto: 380 }),
        fila(S.map(w => celda([
          parrafo(texto(''), { antes: 500 }),
          lineaFirma('Firma:'), lineaFirma('Nombre:'), lineaFirma('Cargo:'),
          parrafo(texto(''), { despues: 60 })
        ], w, { vAlign: VerticalAlign.BOTTOM })), { alto: 1900 })
      ])
    ];

    return new Document({
      creator: 'Hub Comex',
      title: `${esTraslado ? 'Informe de traslado' : 'Informe de despacho'} - ${reg.nombre}`,
      styles: { default: { document: { run: { font: 'Aptos', size: 18, color: COLOR.texto } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1660, right: 936, bottom: 1000, left: 936, header: 500, footer: 450 } } },
        headers: { default: encabezado },
        footers: { default: pie },
        children: [...datos, ...transporte, ...inspeccion, ...carga, ...sellos, ...observaciones, ...fotos, ...firmas]
      }]
    });
  }

  const nombreArchivo = reg => `${reg.tipo === 'traslado' ? 'Informe de traslado' : 'Informe de despacho'} - ${reg.carpeta || reg.nombre}.docx`
    .replace(/[\\/:*?"<>|]/g, '-');

  const api = { crear, nombreArchivo, INSPECCION };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  raiz.HCInforme = api;
})(typeof window !== 'undefined' ? window : globalThis);
