const ExcelJS = require('exceljs');
const { aTexto, normalizarTexto } = require('./textUtils');
const {
  CATEGORIAS,
  CATEGORIA_DETALLE,
  ORDEN_CATEGORIA
} = require('./classifier');

const COLUMNAS_SALIDA = [
  { key: 'ubicacion', label: 'Ubicación origen', width: 21, align: 'left' },
  { key: 'sucursal', label: 'Sucursal', width: 23, align: 'left' },
  { key: 'mod', label: 'Mod', width: 8, align: 'center' },
  { key: 'descripcion', label: 'Descripción', width: 54, align: 'left' },
  { key: 'color', label: 'Color', width: 22, align: 'left' },
  { key: 'serie', label: 'Serie', width: 16, align: 'left' },
  { key: 'dealer', label: 'Dealer', width: 24, align: 'left' },
  { key: 'dias', label: 'Días', width: 9, align: 'right', numFmt: '0' },
  { key: 'clasificacion', label: 'Clasificación', width: 14, align: 'center' },
  { key: 'categoria', label: 'Categoría', width: 17, align: 'left' },
  { key: 'observaciones', label: 'Observaciones', width: 46, align: 'left' }
];

const PALETA = {
  negro: 'FF09090B',
  carbon: 'FF14151A',
  grisOscuro: 'FF1D1E24',
  grisMedio: 'FF4B5563',
  gris: 'FF687182',
  grisSuave: 'FF8B93A1',
  grisClaro: 'FFF3F4F7',
  grisFondo: 'FFF8F9FB',
  borde: 'FFE2E5EA',
  bordeSuave: 'FFEDEFF3',
  blanco: 'FFFFFFFF',
  rojo: 'FFC3002F',
  rojoOscuro: 'FF940022',
  rojoSuave: 'FFFFE9EF',
  verde: 'FF0F7B53',
  verdeSuave: 'FFE7F6EE',
  azul: 'FF2563EB',
  azulSuave: 'FFE8EFFF',
  amber: 'FFB45309',
  amberSuave: 'FFFFF1E0'
};

const DESCRIPCION_CATEGORIA = {
  Importados: 'Magnite, X-Trail, Kait, Urvan, Pathfinder, Z, Kicks Play e-Power y Frontier V6 / Híbrido (versiones importadas).',
  Compactos: 'Versa, Sentra, Kicks regular y V-Drive.',
  NP300: 'NP300 y Frontier de 4 cilindros (versiones domésticas).',
  Subcompactos: 'March.',
  'Sin clasificar': 'Modelos que no coinciden con el catálogo de clasificación.'
};

function estiloBase(extra = {}) {
  return {
    font: {
      name: 'Calibri',
      size: 10,
      color: { argb: PALETA.grisOscuro },
      ...(extra.font || {})
    },
    alignment: {
      vertical: 'middle',
      ...(extra.alignment || {})
    },
    ...(extra.fill ? { fill: extra.fill } : {}),
    ...(extra.border ? { border: extra.border } : {}),
    ...(extra.numFmt ? { numFmt: extra.numFmt } : {})
  };
}

function aplicarEstilo(celda, extra = {}) {
  const base = estiloBase(extra);

  if (base.font) celda.font = base.font;
  if (base.alignment) celda.alignment = base.alignment;
  if (base.fill) celda.fill = base.fill;
  if (base.border) celda.border = base.border;
  if (base.numFmt) celda.numFmt = base.numFmt;
}

function fondo(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function aplicarFondo(hoja, fila1, columna1, fila2, columna2, argb) {
  for (let fila = fila1; fila <= fila2; fila += 1) {
    for (let columna = columna1; columna <= columna2; columna += 1) {
      hoja.getCell(fila, columna).fill = fondo(argb);
    }
  }
}

function aplicarBorde(hoja, fila1, columna1, fila2, columna2, lado, color, estilo = 'thin') {
  for (let fila = fila1; fila <= fila2; fila += 1) {
    for (let columna = columna1; columna <= columna2; columna += 1) {
      const celda = hoja.getCell(fila, columna);
      celda.border = {
        ...(celda.border || {}),
        [lado]: { style: estilo, color: { argb: color } }
      };
    }
  }
}

function combinar(hoja, fila1, columna1, fila2, columna2, valor, extra = {}) {
  if (fila1 !== fila2 || columna1 !== columna2) {
    hoja.mergeCells(fila1, columna1, fila2, columna2);
  }

  const celda = hoja.getCell(fila1, columna1);
  if (valor !== undefined && valor !== null) celda.value = valor;
  aplicarEstilo(celda, extra);
  return celda;
}

function numero(valor) {
  return Number.isFinite(Number(valor)) ? Number(valor) : 0;
}

function porcentaje(valor, total) {
  if (!total) return 0;
  return (valor / total) * 100;
}

function formatearPorcentaje(valor) {
  return `${valor.toFixed(1)}%`;
}

function contarPor(registros, campo) {
  const mapa = new Map();

  for (const registro of registros) {
    const valor = aTexto(registro[campo]) || 'Sin dato';
    mapa.set(valor, (mapa.get(valor) || 0) + 1);
  }

  return [...mapa.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      return a.label.localeCompare(b.label, 'es', { numeric: true });
    });
}

function centrosCategoria(registros) {
  const mapa = new Map();

  for (const registro of registros) {
    if (!mapa.has(registro.ubicacion)) {
      mapa.set(registro.ubicacion, {
        total: 0,
        Importados: 0,
        Compactos: 0,
        NP300: 0,
        Subcompactos: 0,
        'Sin clasificar': 0
      });
    }

    const fila = mapa.get(registro.ubicacion);
    fila.total += 1;
    fila[registro.categoria] = (fila[registro.categoria] || 0) + 1;
  }

  return mapa;
}

function crearTarjetas(hoja, filaInicio, items, tema) {
  const rangos = [
    [1, 2],
    [3, 4],
    [5, 6],
    [7, 8],
    [9, 10]
  ];

  items.forEach((item, indice) => {
    const rango = rangos[indice];
    if (!rango) return;
    const [columna1, columna2] = rango;
    const esOscuro = tema === 'oscuro';

    aplicarFondo(hoja, filaInicio, columna1, filaInicio + 2, columna2, item.fondo);
    if (item.colorBorde) {
      aplicarBorde(hoja, filaInicio, columna1, filaInicio, columna2, 'top', item.colorBorde, 'medium');
    }

    combinar(hoja, filaInicio, columna1, filaInicio, columna2, item.label, {
      font: {
        size: 8,
        bold: true,
        color: { argb: esOscuro ? 'FFB9C0CA' : PALETA.gris }
      },
      alignment: { horizontal: 'center', vertical: 'middle' }
    });

    combinar(hoja, filaInicio + 1, columna1, filaInicio + 1, columna2, item.valor, {
      font: {
        size: 22,
        bold: true,
        color: { argb: item.colorValor }
      },
      alignment: { horizontal: 'center', vertical: 'middle' },
      numFmt: '0'
    });

    combinar(hoja, filaInicio + 2, columna1, filaInicio + 2, columna2, item.sub, {
      font: {
        size: 8,
        bold: true,
        color: { argb: item.colorSub }
      },
      alignment: { horizontal: 'center', vertical: 'middle' }
    });
  });
}

function crearEtiquetaSeccion(hoja, fila, texto, tema = 'claro') {
  const esOscuro = tema === 'oscuro';
  const celda = combinar(hoja, fila, 1, fila, 10, texto, {
    font: {
      size: 9,
      bold: true,
      color: { argb: esOscuro ? 'FFFF7A90' : PALETA.rojo }
    },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 }
  });

  hoja.getRow(fila).height = 22;
  return celda;
}

function crearHojaResumen(libro, datos) {
  const hoja = libro.addWorksheet('RESUMEN', {
    views: [{ showGridLines: false }],
    properties: { tabColor: { argb: PALETA.grisOscuro } }
  });

  const anchos = [30, 14, 14, 14, 14, 14, 14, 14, 14, 14];
  anchos.forEach((ancho, indice) => {
    hoja.getColumn(indice + 1).width = ancho;
  });

  const resumen = datos.resumen;
  const registros = datos.registros;
  const total = numero(resumen.totalRegistros);
  const conteos = new Map((resumen.porCategoria || []).map((item) => [item.label, numero(item.count)]));

  hoja.getRow(1).height = 44;
  combinar(hoja, 1, 1, 1, 10, 'LISTA DE EXISTENCIAS', {
    font: { size: 22, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon)
  });

  hoja.getRow(2).height = 5;
  aplicarFondo(hoja, 2, 1, 2, 10, PALETA.rojo);

  hoja.getRow(3).height = 24;
  combinar(hoja, 3, 1, 3, 7, 'Resumen ejecutivo · Inventario de vehículos Nissan', {
    font: { size: 11, bold: true, color: { argb: PALETA.grisOscuro } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 }
  });
  combinar(hoja, 3, 8, 3, 10, 'USO INTERNO', {
    font: { size: 8, bold: true, color: { argb: PALETA.gris } },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 }
  });

  const corte = resumen.fechaCorte || 'No detectada';
  hoja.getRow(4).height = 18;
  combinar(hoja, 4, 1, 4, 10, `Fuente: ${datos.nombreArchivo}   ·   Corte: ${corte}   ·   Generado: ${datos.fechaGeneracion}`, {
    font: { size: 9, color: { argb: PALETA.gris } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 }
  });

  const tarjetas = [
    {
      label: 'TOTAL GLOBAL',
      valor: total,
      sub: 'Unidades en el reporte',
      fondo: PALETA.rojo,
      colorValor: PALETA.blanco,
      colorSub: 'FFFFD7DF'
    },
    ...CATEGORIAS.map((categoria) => {
      const detalle = CATEGORIA_DETALLE[categoria];
      const count = conteos.get(categoria) || 0;
      return {
        label: categoria.toUpperCase(),
        valor: count,
        sub: `${detalle.letra} · ${formatearPorcentaje(porcentaje(count, total))}`,
        fondo: detalle.fondo,
        colorValor: detalle.color,
        colorSub: PALETA.gris,
        colorBorde: detalle.color
      };
    })
  ];

  hoja.getRow(6).height = 18;
  hoja.getRow(7).height = 32;
  hoja.getRow(8).height = 16;
  crearTarjetas(hoja, 6, tarjetas, 'claro');

  crearEtiquetaSeccion(hoja, 10, 'DISTRIBUCIÓN POR CATEGORÍA');

  const filaCategorias = 11;
  const columnasCategorias = [
    ['Categoría', 1, 1],
    ['Clave', 2, 2],
    ['Unidades', 3, 3],
    ['% del total', 4, 4],
    ['Descripción', 5, 10]
  ];

  hoja.getRow(filaCategorias).height = 24;
  for (const [label, columna1, columna2] of columnasCategorias) {
    combinar(hoja, filaCategorias, columna1, filaCategorias, columna2, label, {
      font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
      alignment: { horizontal: columna1 === 1 ? 'left' : 'center', vertical: 'middle', indent: columna1 === 1 ? 1 : 0 },
      fill: fondo(PALETA.grisOscuro)
    });
  }
  aplicarBorde(hoja, filaCategorias, 1, filaCategorias, 10, 'bottom', PALETA.rojo, 'medium');

  const filasCategoria = [];
  for (const categoria of CATEGORIAS) {
    const detalle = CATEGORIA_DETALLE[categoria];
    const count = conteos.get(categoria) || 0;
    filasCategoria.push({ categoria, detalle, count });
  }
  if ((conteos.get('Sin clasificar') || 0) > 0) {
    filasCategoria.push({
      categoria: 'Sin clasificar',
      detalle: CATEGORIA_DETALLE['Sin clasificar'],
      count: conteos.get('Sin clasificar')
    });
  }

  filasCategoria.forEach((item, indice) => {
    const fila = filaCategorias + 1 + indice;
    hoja.getRow(fila).height = 24;
    combinar(hoja, fila, 1, fila, 1, item.categoria, {
      font: { size: 10, bold: true, color: { argb: item.detalle.color } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
      fill: fondo(item.detalle.fondo)
    });
    combinar(hoja, fila, 2, fila, 2, item.detalle.letra, {
      font: { size: 10, bold: true, color: { argb: PALETA.blanco } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      fill: fondo(item.detalle.color)
    });
    combinar(hoja, fila, 3, fila, 3, item.count, {
      font: { size: 10, bold: true, color: { argb: PALETA.grisOscuro } },
      alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
      numFmt: '0'
    });
    combinar(hoja, fila, 4, fila, 4, porcentaje(item.count, total) / 100, {
      font: { size: 10, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
      numFmt: '0.0%'
    });
    combinar(hoja, fila, 5, fila, 10, DESCRIPCION_CATEGORIA[item.categoria] || '', {
      font: { size: 9, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1, wrapText: true }
    });
    aplicarBorde(hoja, fila, 1, fila, 10, 'bottom', PALETA.bordeSuave);
  });

  const filaTotalCategorias = filaCategorias + 1 + filasCategoria.length;
  hoja.getRow(filaTotalCategorias).height = 26;
  combinar(hoja, filaTotalCategorias, 1, filaTotalCategorias, 2, 'TOTAL GLOBAL', {
    font: { size: 10, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon)
  });
  combinar(hoja, filaTotalCategorias, 3, filaTotalCategorias, 3, total, {
    font: { size: 11, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon),
    numFmt: '0'
  });
  combinar(hoja, filaTotalCategorias, 4, filaTotalCategorias, 4, 1, {
    font: { size: 10, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon),
    numFmt: '0.0%'
  });
  combinar(hoja, filaTotalCategorias, 5, filaTotalCategorias, 10, 'Total de unidades del reporte', {
    font: { size: 9, color: { argb: 'FFD7DBE2' } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon)
  });

  const rangoBarrasCategoria = `C${filaCategorias + 1}:C${filaTotalCategorias - 1}`;
  if (filasCategoria.length) {
    hoja.addConditionalFormatting({
      ref: rangoBarrasCategoria,
      rules: [{
        type: 'dataBar',
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: { argb: 'FFF6C6CF' }
      }]
    });
  }

  const filaUbicacionesTitulo = filaTotalCategorias + 2;
  crearEtiquetaSeccion(hoja, filaUbicacionesTitulo, 'DISTRIBUCIÓN POR UBICACIÓN DE ORIGEN');

  const filaUbicaciones = filaUbicacionesTitulo + 1;
  const haySinClasificar = (conteos.get('Sin clasificar') || 0) > 0;
  const encabezadosUbicacion = [
    'Ubicación',
    'Unidades',
    '%',
    'Importados',
    'Compactos',
    'NP300',
    'Subcompactos',
    ...(haySinClasificar ? ['Sin clasificar'] : [])
  ];

  hoja.getRow(filaUbicaciones).height = 24;
  encabezadosUbicacion.forEach((label, indice) => {
    combinar(hoja, filaUbicaciones, indice + 1, filaUbicaciones, indice + 1, label, {
      font: { size: 8, bold: true, color: { argb: PALETA.blanco } },
      alignment: { horizontal: indice === 0 ? 'left' : 'center', vertical: 'middle', indent: indice === 0 ? 1 : 0, wrapText: true },
      fill: fondo(PALETA.grisOscuro)
    });
  });
  aplicarBorde(hoja, filaUbicaciones, 1, filaUbicaciones, encabezadosUbicacion.length, 'bottom', PALETA.rojo, 'medium');

  const centros = centrosCategoria(registros);
  const ubicaciones = (resumen.porUbicacion || []).map((item) => ({
    label: item.label,
    count: numero(item.count),
    centros: centros.get(item.label) || {
      total: 0,
      Importados: 0,
      Compactos: 0,
      NP300: 0,
      Subcompactos: 0,
      'Sin clasificar': 0
    }
  }));

  ubicaciones.forEach((item, indice) => {
    const fila = filaUbicaciones + 1 + indice;
    hoja.getRow(fila).height = 22;
    const valores = [
      item.label,
      item.count,
      porcentaje(item.count, total) / 100,
      item.centros.Importados || 0,
      item.centros.Compactos || 0,
      item.centros.NP300 || 0,
      item.centros.Subcompactos || 0,
      ...(haySinClasificar ? [item.centros['Sin clasificar'] || 0] : [])
    ];

    valores.forEach((valor, columna) => {
      combinar(hoja, fila, columna + 1, fila, columna + 1, valor, {
        font: {
          size: 9,
          bold: columna <= 1,
          color: { argb: columna === 0 ? PALETA.grisOscuro : PALETA.gris }
        },
        alignment: {
          horizontal: columna === 0 ? 'left' : 'center',
          vertical: 'middle',
          indent: columna === 0 ? 1 : 0
        },
        numFmt: columna === 2 ? '0.0%' : (columna === 1 || columna >= 3 ? '0' : undefined)
      });
    });
    aplicarBorde(hoja, fila, 1, fila, valores.length, 'bottom', PALETA.bordeSuave);
  });

  const filaTotalUbicaciones = filaUbicaciones + 1 + ubicaciones.length;
  hoja.getRow(filaTotalUbicaciones).height = 24;
  const totalesUbicacion = [
    'TOTAL GLOBAL',
    total,
    1,
    conteos.get('Importados') || 0,
    conteos.get('Compactos') || 0,
    conteos.get('NP300') || 0,
    conteos.get('Subcompactos') || 0,
    ...(haySinClasificar ? [conteos.get('Sin clasificar') || 0] : [])
  ];
  totalesUbicacion.forEach((valor, columna) => {
    combinar(hoja, filaTotalUbicaciones, columna + 1, filaTotalUbicaciones, columna + 1, valor, {
      font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
      alignment: {
        horizontal: columna === 0 ? 'left' : 'center',
        vertical: 'middle',
        indent: columna === 0 ? 1 : 0
      },
      fill: fondo(PALETA.carbon),
      numFmt: columna === 2 ? '0.0%' : '0'
    });
  });

  if (ubicaciones.length) {
    hoja.addConditionalFormatting({
      ref: `B${filaUbicaciones + 1}:B${filaTotalUbicaciones - 1}`,
      rules: [{
        type: 'dataBar',
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: { argb: 'FFC7D8FF' }
      }]
    });
  }

  const filaModelosTitulo = filaTotalUbicaciones + 2;
  crearEtiquetaSeccion(hoja, filaModelosTitulo, 'DISTRIBUCIÓN POR MODELO');

  const filaModelos = filaModelosTitulo + 1;
  ['Modelo', 'Categoría', 'Clave', 'Unidades', '%'].forEach((label, indice) => {
    combinar(hoja, filaModelos, indice + 1, filaModelos, indice + 1, label, {
      font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
      alignment: { horizontal: indice === 0 ? 'left' : 'center', vertical: 'middle', indent: indice === 0 ? 1 : 0 },
      fill: fondo(PALETA.grisOscuro)
    });
  });
  aplicarBorde(hoja, filaModelos, 1, filaModelos, 5, 'bottom', PALETA.rojo, 'medium');

  const gruposModelo = new Map();

  for (const registro of registros) {
    const clave = `${registro.modelo}\u0000${registro.categoria}`;
    gruposModelo.set(clave, (gruposModelo.get(clave) || 0) + 1);
  }

  const modelos = [...gruposModelo.entries()]
    .map(([clave, count]) => {
      const [modelo, categoria] = clave.split('\u0000');
      return { modelo, categoria, count };
    })
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      const ordenA = ORDEN_CATEGORIA.get(a.categoria) ?? 99;
      const ordenB = ORDEN_CATEGORIA.get(b.categoria) ?? 99;
      if (ordenA !== ordenB) return ordenA - ordenB;
      return a.modelo.localeCompare(b.modelo, 'es', { numeric: true });
    });

  modelos.forEach((item, indice) => {
    const fila = filaModelos + 1 + indice;
    const detalle = CATEGORIA_DETALLE[item.categoria] || CATEGORIA_DETALLE['Sin clasificar'];
    hoja.getRow(fila).height = 21;
    combinar(hoja, fila, 1, fila, 1, item.modelo, {
      font: { size: 9, bold: true, color: { argb: PALETA.grisOscuro } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1 }
    });
    combinar(hoja, fila, 2, fila, 2, item.categoria, {
      font: { size: 9, bold: true, color: { argb: detalle.color } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
      fill: fondo(detalle.fondo)
    });
    combinar(hoja, fila, 3, fila, 3, detalle.letra, {
      font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      fill: fondo(detalle.color)
    });
    combinar(hoja, fila, 4, fila, 4, item.count, {
      font: { size: 9, bold: true, color: { argb: PALETA.grisOscuro } },
      alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
      numFmt: '0'
    });
    combinar(hoja, fila, 5, fila, 5, porcentaje(item.count, total) / 100, {
      font: { size: 9, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
      numFmt: '0.0%'
    });
    aplicarBorde(hoja, fila, 1, fila, 5, 'bottom', PALETA.bordeSuave);
  });

  const filaTotalModelos = filaModelos + 1 + modelos.length;
  hoja.getRow(filaTotalModelos).height = 24;
  ['TOTAL GLOBAL', '', '', total, 1].forEach((valor, indice) => {
    combinar(hoja, filaTotalModelos, indice + 1, filaTotalModelos, indice + 1, valor, {
      font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
      alignment: { horizontal: indice === 0 ? 'left' : 'center', vertical: 'middle', indent: indice === 0 ? 1 : 0 },
      fill: fondo(PALETA.carbon),
      numFmt: indice === 4 ? '0.0%' : '0'
    });
  });

  if (modelos.length) {
    hoja.addConditionalFormatting({
      ref: `D${filaModelos + 1}:D${filaTotalModelos - 1}`,
      rules: [{
        type: 'dataBar',
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: { argb: 'FFC6E9D8' }
      }]
    });
  }

  const filaSeccionesTitulo = filaTotalModelos + 2;
  crearEtiquetaSeccion(hoja, filaSeccionesTitulo, 'CONTROL DEL ARCHIVO ORIGEN');

  const filaSecciones = filaSeccionesTitulo + 1;
  hoja.getRow(filaSecciones).height = 24;
  combinar(hoja, filaSecciones, 1, filaSecciones, 1, 'Sección', {
    font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.grisOscuro)
  });
  combinar(hoja, filaSecciones, 2, filaSecciones, 2, 'Estado', {
    font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'center', vertical: 'middle' },
    fill: fondo(PALETA.grisOscuro)
  });
  combinar(hoja, filaSecciones, 3, filaSecciones, 3, 'Unidades', {
    font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.grisOscuro)
  });
  combinar(hoja, filaSecciones, 4, filaSecciones, 10, 'Nota', {
    font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.grisOscuro)
  });
  aplicarBorde(hoja, filaSecciones, 1, filaSecciones, 10, 'bottom', PALETA.rojo, 'medium');

  const secciones = resumen.secciones || [];
  secciones.forEach((seccion, indice) => {
    const fila = filaSecciones + 1 + indice;
    const color = seccion.incluida ? PALETA.verde : PALETA.rojo;
    const fondoFila = seccion.incluida ? PALETA.verdeSuave : PALETA.rojoSuave;
    hoja.getRow(fila).height = 22;
    combinar(hoja, fila, 1, fila, 1, seccion.label, {
      font: { size: 9, bold: true, color: { argb: color } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
      fill: fondo(fondoFila)
    });
    combinar(hoja, fila, 2, fila, 2, seccion.incluida ? 'Incluida' : 'Excluida', {
      font: { size: 9, bold: true, color: { argb: color } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      fill: fondo(fondoFila)
    });
    combinar(hoja, fila, 3, fila, 3, seccion.count, {
      font: { size: 9, bold: true, color: { argb: PALETA.grisOscuro } },
      alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
      numFmt: '0'
    });
    combinar(hoja, fila, 4, fila, 10, seccion.incluida
      ? 'Se conservó en la lista de existencias.'
      : 'Se excluyó del reporte por pertenecer a una sección no incluida.', {
      font: { size: 9, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1 }
    });
    aplicarBorde(hoja, fila, 1, fila, 10, 'bottom', PALETA.bordeSuave);
  });

  const filaTotalSecciones = filaSecciones + 1 + secciones.length;
  hoja.getRow(filaTotalSecciones).height = 24;
  combinar(hoja, filaTotalSecciones, 1, filaTotalSecciones, 2, 'TOTAL DE UNIDADES', {
    font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon)
  });
  combinar(hoja, filaTotalSecciones, 3, filaTotalSecciones, 3, total + numero(resumen.totalExcluidos), {
    font: { size: 10, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon),
    numFmt: '0'
  });
  combinar(hoja, filaTotalSecciones, 4, filaTotalSecciones, 10, `${total} incluidas · ${numero(resumen.totalExcluidos)} excluidas`, {
    font: { size: 9, color: { argb: 'FFD7DBE2' } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(PALETA.carbon)
  });

  const filaNotasTitulo = filaTotalSecciones + 2;
  crearEtiquetaSeccion(hoja, filaNotasTitulo, 'NOTAS DEL PROCESAMIENTO');

  const notas = [...(datos.notasInterpretacion || []), ...(datos.advertencias || [])];
  const notasFinales = notas.length ? notas.slice(0, 12) : ['Sin observaciones adicionales.'];

  notasFinales.forEach((nota, indice) => {
    const fila = filaNotasTitulo + 1 + indice;
    hoja.getRow(fila).height = 24;
    combinar(hoja, fila, 1, fila, 10, `•  ${nota}`, {
      font: { size: 9, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1, wrapText: true }
    });
  });

  const ultimaFila = filaNotasTitulo + notasFinales.length + 1;
  hoja.pageSetup = {
    paperSize: 9,
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
    printArea: `A1:J${ultimaFila}`
  };

  hoja.headerFooter.oddFooter = '&L&8Organizador de Inventarios Nissan&C&8Resumen ejecutivo&R&8Página &P de &N';

  return hoja;
}

function crearHojaDatos(libro, nombreHoja, registros, opciones) {
  const {
    titulo,
    color,
    letra,
    totalEtiqueta
  } = opciones;

  const hoja = libro.addWorksheet(nombreHoja, {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 2 }],
    properties: { tabColor: { argb: color } }
  });

  COLUMNAS_SALIDA.forEach((columna, indice) => {
    hoja.getColumn(indice + 1).width = columna.width;
  });

  hoja.getRow(1).height = 32;
  const tituloCelda = combinar(hoja, 1, 1, 1, COLUMNAS_SALIDA.length, null, {
    alignment: { vertical: 'middle', horizontal: 'left', indent: 1 },
    fill: fondo(color)
  });
  tituloCelda.value = {
    richText: [
      { text: titulo, font: { name: 'Calibri', size: 14, bold: true, color: { argb: PALETA.blanco } } },
      {
        text: `  ·  ${registros.length} unidad(es)  ·  ordenado por antigüedad`,
        font: { name: 'Calibri', size: 9, color: { argb: 'FFE4E7EC' } }
      }
    ]
  };

  hoja.getRow(2).height = 32;
  COLUMNAS_SALIDA.forEach((columna, indice) => {
    const celda = combinar(hoja, 2, indice + 1, 2, indice + 1, columna.label, {
      font: { size: 9, bold: true, color: { argb: PALETA.blanco } },
      alignment: {
        horizontal: columna.align === 'right' ? 'right' : (columna.align === 'center' ? 'center' : 'left'),
        vertical: 'middle',
        wrapText: true,
        indent: columna.align === 'left' ? 1 : 0
      },
      fill: fondo(PALETA.grisOscuro)
    });
    celda.border = {
      bottom: { style: 'medium', color: { argb: color } }
    };
  });

  if (!registros.length) {
    hoja.getRow(3).height = 26;
    combinar(hoja, 3, 1, 3, COLUMNAS_SALIDA.length, 'Sin unidades en esta categoría.', {
      font: { size: 10, italic: true, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'center', vertical: 'middle' }
    });
  }

  registros.forEach((registro, indice) => {
    const fila = 3 + indice;
    const filaHoja = hoja.getRow(fila);
    filaHoja.height = 19;

    COLUMNAS_SALIDA.forEach((columna, columnaIndice) => {
      const celda = filaHoja.getCell(columnaIndice + 1);
      const valor = registro[columna.key];
      celda.value = valor === undefined || valor === null ? '' : valor;

      const esImpar = indice % 2 === 1;
      const colorCategoria = CATEGORIA_DETALLE[registro.categoria] || CATEGORIA_DETALLE['Sin clasificar'];
      const esCategoria = columna.key === 'categoria';
      const esClasificacion = columna.key === 'clasificacion';
      const esDias = columna.key === 'dias';

      let fondoFila = esImpar ? PALETA.grisFondo : PALETA.blanco;
      let colorTexto = PALETA.grisOscuro;
      let negrita = false;
      let alineacion = columna.align || 'left';

      if (esCategoria) {
        fondoFila = colorCategoria.fondo;
        colorTexto = colorCategoria.color;
        negrita = true;
      }

      if (esClasificacion) {
        fondoFila = colorCategoria.color;
        colorTexto = PALETA.blanco;
        negrita = true;
        alineacion = 'center';
      }

      if (esDias) {
        alineacion = 'right';
        if (numero(valor) >= 365) {
          colorTexto = PALETA.rojoOscuro;
          negrita = true;
        } else if (numero(valor) >= 180) {
          colorTexto = PALETA.amber;
          negrita = true;
        }
      }

      if (columna.key === 'observaciones') {
        colorTexto = PALETA.gris;
      }

      celda.font = {
        name: 'Calibri',
        size: esCategoria || esClasificacion ? 9 : 10,
        bold: negrita,
        color: { argb: colorTexto }
      };
      celda.fill = fondo(fondoFila);
      celda.alignment = {
        horizontal: alineacion,
        vertical: 'middle',
        indent: alineacion === 'left' ? 1 : 0
      };
      if (columna.numFmt) celda.numFmt = columna.numFmt;
      if (columna.key === 'serie' || columna.key === 'dealer' || columna.key === 'mod') {
        celda.numFmt = '@';
      }
      celda.border = {
        bottom: { style: 'hair', color: { argb: PALETA.bordeSuave } }
      };
    });
  });

  const ultimaFilaDatos = 2 + registros.length;
  if (registros.length) {
    hoja.autoFilter = {
      from: { row: 2, column: 1 },
      to: { row: ultimaFilaDatos, column: COLUMNAS_SALIDA.length }
    };

    hoja.addConditionalFormatting({
      ref: `H3:H${ultimaFilaDatos}`,
      rules: [{
        type: 'dataBar',
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: { argb: 'FFF6C6CF' }
      }]
    });
  }

  const filaTotal = ultimaFilaDatos + 2;
  hoja.getRow(filaTotal).height = 30;
  combinar(hoja, filaTotal, 1, filaTotal, COLUMNAS_SALIDA.length, `${totalEtiqueta}  ·  ${registros.length} unidad(es)`, {
    font: { size: 11, bold: true, color: { argb: PALETA.blanco } },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    fill: fondo(color)
  });

  if (letra) {
    const filaLetra = filaTotal + 1;
    hoja.getRow(filaLetra).height = 18;
    combinar(hoja, filaLetra, 1, filaLetra, COLUMNAS_SALIDA.length, `Clave de categoría: ${letra}`, {
      font: { size: 8, color: { argb: PALETA.gris } },
      alignment: { horizontal: 'left', vertical: 'middle', indent: 1 }
    });
  }

  const ultimaFila = filaTotal + (letra ? 1 : 0);

  hoja.properties.defaultRowHeight = 18;
  hoja.pageSetup = {
    paperSize: 9,
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    printTitlesRow: '1:2',
    margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
    printArea: `A1:K${ultimaFila}`
  };

  hoja.headerFooter.oddHeader = '&L&8Organizador de Inventarios Nissan&C&8Lista de existencias&R&8' + titulo;
  hoja.headerFooter.oddFooter = '&L&8Reporte generado automáticamente&C&8&Página &P de &N&R&8El original no se modifica';

  return hoja;
}

function nombreHojaSeguro(texto, usados) {
  let base = String(texto ?? '')
    .replace(/[\\/?*[\]:]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^'+|'+$/g, '')
    .trim()
    .slice(0, 31);

  if (!base) base = 'Hoja';

  let nombre = base;
  let contador = 2;

  while (usados.has(normalizarTexto(nombre))) {
    const sufijo = ` (${contador})`;
    nombre = `${base.slice(0, 31 - sufijo.length)}${sufijo}`;
    contador += 1;
  }

  usados.add(normalizarTexto(nombre));
  return nombre;
}

function agruparPorCategoria(registros) {
  const grupos = new Map();

  for (const categoria of CATEGORIAS) {
    grupos.set(categoria, registros.filter((registro) => registro.categoria === categoria));
  }

  const sinClasificar = registros.filter((registro) => registro.categoria === 'Sin clasificar');
  return { grupos, sinClasificar };
}

async function construirLibroSalida(datos) {
  const libro = new ExcelJS.Workbook();
  libro.creator = 'Organizador de Inventarios Nissan';
  libro.lastModifiedBy = 'Organizador de Inventarios Nissan';
  libro.created = new Date();
  libro.modified = new Date();
  libro.title = 'Lista de existencias Nissan';
  libro.subject = 'Inventario segmentado por categoría y ubicación';
  libro.description = 'Reporte ejecutivo generado automáticamente a partir de la lista de existencias.';

  const nombresUsados = new Set();

  crearHojaResumen(libro, datos);

  const nombreLista = nombreHojaSeguro('LISTA DE EXISTENCIAS', nombresUsados);
  crearHojaDatos(libro, nombreLista, datos.registros, {
    titulo: 'LISTA DE EXISTENCIAS',
    color: PALETA.carbon,
    letra: '',
    totalEtiqueta: 'TOTAL GLOBAL'
  });

  const { grupos, sinClasificar } = agruparPorCategoria(datos.registros);

  for (const categoria of CATEGORIAS) {
    const detalle = CATEGORIA_DETALLE[categoria];
    const nombre = nombreHojaSeguro(categoria.toUpperCase(), nombresUsados);
    crearHojaDatos(libro, nombre, grupos.get(categoria) || [], {
      titulo: categoria.toUpperCase(),
      color: detalle.color,
      letra: detalle.letra,
      totalEtiqueta: `TOTAL ${categoria.toUpperCase()}`
    });
  }

  if (sinClasificar.length) {
    const nombre = nombreHojaSeguro('SIN CLASIFICAR', nombresUsados);
    crearHojaDatos(libro, nombre, sinClasificar, {
      titulo: 'SIN CLASIFICAR',
      color: CATEGORIA_DETALLE['Sin clasificar'].color,
      letra: '?',
      totalEtiqueta: 'TOTAL SIN CLASIFICAR'
    });
  }

  for (const ubicacion of datos.resumen.porUbicacion || []) {
    const registrosUbicacion = datos.registros.filter(
      (registro) => registro.ubicacion === ubicacion.label
    );

    if (!registrosUbicacion.length) continue;

    const nombre = nombreHojaSeguro(ubicacion.label.toUpperCase(), nombresUsados);
    crearHojaDatos(libro, nombre, registrosUbicacion, {
      titulo: `UBICACIÓN · ${ubicacion.label.toUpperCase()}`,
      color: PALETA.azul,
      letra: '',
      totalEtiqueta: `TOTAL ${ubicacion.label.toUpperCase()}`
    });
  }

  const resultado = await libro.xlsx.writeBuffer();
  return Buffer.from(resultado);
}

module.exports = {
  COLUMNAS_SALIDA,
  construirLibroSalida,
  construirLibroSalidaFormateado: construirLibroSalida,
  nombreHojaSeguro,
  contarPor,
  agruparPorCategoria
};
