const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { normalizarTexto, aTexto } = require('./textUtils');

const COLUMNAS_DERIVADAS = [
  'Hoja de origen',
  'Fila original',
  'Tipo identificado',
  'Sucursal identificada',
  'Procedencia identificada',
  'Revisión'
];

const colador = new Intl.Collator('es', { sensitivity: 'base', numeric: true });

function contarValores(registros, campo) {
  const mapa = new Map();

  for (const registro of registros) {
    const valor = aTexto(registro[campo]);
    const clave = normalizarTexto(valor) || '__sin_clasificar__';
    const label = valor || 'Sin clasificar';

    if (!mapa.has(clave)) {
      mapa.set(clave, { label, count: 0 });
    }

    mapa.get(clave).count += 1;
  }

  return [...mapa.values()].sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return colador.compare(a.label, b.label);
  });
}

function contarUnicos(registros, campo) {
  const valores = new Set();

  for (const registro of registros) {
    const valor = normalizarTexto(registro[campo]);
    if (valor) valores.add(valor);
  }

  return valores.size;
}

function agregarColumnaUnion(columnas, mapa, key, label) {
  if (mapa.has(key)) return;
  const columna = { key, label };
  mapa.set(key, columna);
  columnas.push(columna);
}

function crearMapaUnion(hojas) {
  const columnas = [];
  const mapa = new Map();
  const mapasPorHoja = new Map();

  for (const hoja of hojas) {
    if (hoja.vacia) continue;

    const mapaHoja = new Map();

    for (const encabezado of hoja.encabezados) {
      let key;
      let label;

      if (encabezado.generada) {
        key = `__${normalizarTexto(hoja.hoja) || 'hoja'}__${encabezado.index}`;
        label = `${encabezado.label} (${hoja.hoja})`;
      } else if (encabezado.ocurrencia > 1) {
        key = `${encabezado.base}#${encabezado.ocurrencia}`;
        label = encabezado.label;
      } else {
        key = encabezado.base;
        label = encabezado.label;
      }

      agregarColumnaUnion(columnas, mapa, key, label);
      mapaHoja.set(encabezado.index, key);
    }

    mapasPorHoja.set(hoja.hoja, mapaHoja);
  }

  return { columnas, mapasPorHoja };
}

function aplicarUnion(hojas, mapasPorHoja) {
  const registros = [];

  for (const hoja of hojas) {
    if (hoja.vacia) continue;
    const mapaHoja = mapasPorHoja.get(hoja.hoja) || new Map();

    for (const registro of hoja.registros) {
      const valores = {};

      hoja.encabezados.forEach((encabezado) => {
        const key = mapaHoja.get(encabezado.index);
        valores[key] = registro.valoresOriginales[encabezado.index] ?? '';
      });

      registros.push({
        hoja: registro.hoja,
        filaOriginal: registro.filaOriginal,
        tipo: registro.tipo,
        sucursal: registro.sucursal,
        procedencia: registro.procedencia,
        revision: registro.revision,
        valores
      });
    }
  }

  return registros;
}

function compararRegistros(a, b, claves) {
  for (const clave of claves) {
    const valorA = a[clave] || '';
    const valorB = b[clave] || '';

    if (!valorA && !valorB) continue;
    if (!valorA) return 1;
    if (!valorB) return -1;

    const comparacion = colador.compare(valorA, valorB);
    if (comparacion !== 0) return comparacion;
  }

  const comparacionHoja = colador.compare(a.hoja, b.hoja);
  if (comparacionHoja !== 0) return comparacionHoja;

  return a.filaOriginal - b.filaOriginal;
}

function crearFilasPrevia(registros, columnas, limite) {
  return registros.slice(0, limite).map((registro) => [
    ...columnas.map((columna) => aTexto(registro.valores[columna.key])),
    registro.hoja,
    registro.filaOriginal,
    registro.tipo,
    registro.sucursal,
    registro.procedencia,
    registro.revision.join('; ')
  ]);
}

function calcularAnchos(encabezados, filas) {
  return encabezados.map((encabezado, indice) => {
    if (encabezado === 'Revisión') return { wch: 42 };
    if (encabezado === 'Hoja de origen') return { wch: 20 };
    if (encabezado === 'Fila original') return { wch: 10 };
    if (encabezado === 'Tipo identificado') return { wch: 22 };
    if (encabezado === 'Sucursal identificada') return { wch: 24 };
    if (encabezado === 'Procedencia identificada') return { wch: 25 };

    let maximo = String(encabezado).length;
    const limite = Math.min(filas.length, 200);

    for (let i = 1; i < limite; i += 1) {
      const largo = aTexto(filas[i][indice]).length;
      if (largo > maximo) maximo = largo;
    }

    return { wch: Math.min(Math.max(10, maximo + 2), 40) };
  });
}

function aplicarFormatoFechas(hoja, filas) {
  if (!hoja['!ref']) return;

  const rango = XLSX.utils.decode_range(hoja['!ref']);

  for (let fila = 1; fila <= rango.e.r; fila += 1) {
    for (let columna = 0; columna <= rango.e.c; columna += 1) {
      const valor = filas[fila] ? filas[fila][columna] : undefined;

      if (!(valor instanceof Date) || Number.isNaN(valor.getTime())) continue;

      const direccion = XLSX.utils.encode_cell({ r: fila, c: columna });
      const celda = hoja[direccion];

      if (celda) {
        const tieneHora = valor.getHours() || valor.getMinutes();
        celda.z = tieneHora ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy';
      }
    }
  }
}

function crearHojaDatos(columnas, registros) {
  const encabezados = [
    ...columnas.map((columna) => columna.label),
    ...COLUMNAS_DERIVADAS
  ];

  const filas = [encabezados];

  for (const registro of registros) {
    filas.push([
      ...columnas.map((columna) => registro.valores[columna.key] ?? ''),
      registro.hoja,
      registro.filaOriginal,
      registro.tipo,
      registro.sucursal,
      registro.procedencia,
      registro.revision.join('; ')
    ]);
  }

  const hoja = XLSX.utils.aoa_to_sheet(filas);
  aplicarFormatoFechas(hoja, filas);
  hoja['!cols'] = calcularAnchos(encabezados, filas);

  if (filas.length > 1) {
    hoja['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: filas.length - 1, c: encabezados.length - 1 }
      })
    };
  }

  return hoja;
}

function construirContenidoResumen(datos) {
  const filas = [];
  const merges = [];

  const agregarTitulo = (texto, columnas = 7) => {
    const indice = filas.length;
    filas.push([texto]);
    if (columnas > 1) {
      merges.push({ s: { r: indice, c: 0 }, e: { r: indice, c: columnas - 1 } });
    }
  };

  const agregarLista = (titulo, lista) => {
    filas.push([]);
    agregarTitulo(titulo, 2);
    filas.push(['Valor', 'Registros']);
    if (!lista.length) {
      filas.push(['Sin datos', 0]);
      return;
    }
    for (const item of lista) {
      filas.push([item.label, item.count]);
    }
  };

  agregarTitulo('REPORTE DE ORGANIZACIÓN DE INVENTARIO NISSAN', 7);
  filas.push([]);
  filas.push(['Archivo original', datos.nombreArchivo]);
  filas.push(['Archivo generado', datos.nombreSalida]);
  filas.push(['Fecha de generación', datos.fechaGeneracion]);
  filas.push(['Total de registros', datos.resumen.totalRegistros]);
  filas.push(['Hojas con datos', datos.resumen.hojasConDatos]);
  filas.push(['Hojas vacías', datos.resumen.hojasVacias]);
  filas.push(['Registros para revisión', datos.resumen.paraRevision]);
  filas.push([]);
  filas.push(['Los datos originales se conservan en sus columnas. Las columnas agregadas al final muestran la clasificación automática y las alertas para revisión. El archivo original no se modificó.']);

  filas.push([]);
  agregarTitulo('HOJAS PROCESADAS', 7);
  filas.push(['Hoja', 'Encabezado en fila', 'Registros', 'Tipo', 'Sucursal', 'Procedencia', 'Columnas sin encabezado']);
  for (const hoja of datos.hojas) {
    filas.push([
      hoja.hoja,
      hoja.vacia ? '—' : (hoja.encabezadoFila ?? 'No detectado'),
      hoja.registros.length,
      hoja.dimensiones.tipo ? hoja.dimensiones.tipo.label : 'No detectada',
      hoja.dimensiones.sucursal ? hoja.dimensiones.sucursal.label : 'No detectada',
      hoja.dimensiones.procedencia ? hoja.dimensiones.procedencia.label : 'No detectada',
      hoja.columnasGeneradas.length
    ]);
  }

  agregarLista('RESUMEN POR TIPO', datos.resumen.porTipo);
  agregarLista('RESUMEN POR SUCURSAL', datos.resumen.porSucursal);
  agregarLista('RESUMEN POR PROCEDENCIA', datos.resumen.porProcedencia);

  filas.push([]);
  agregarTitulo('ADVERTENCIAS Y NOTAS', 7);
  if (!datos.advertencias.length) {
    filas.push(['Sin advertencias.']);
  } else {
    for (const advertencia of datos.advertencias) filas.push([advertencia]);
  }

  if (datos.notasInterpretacion.length) {
    filas.push([]);
    agregarTitulo('INTERPRETACIÓN AUTOMÁTICA', 7);
    for (const nota of datos.notasInterpretacion) filas.push([nota]);
  }

  if (datos.notasPrevias.length) {
    filas.push([]);
    agregarTitulo('NOTAS ENCONTRADAS ANTES DEL ENCABEZADO', 7);
    filas.push(['Hoja', 'Fila', 'Contenido']);
    for (const nota of datos.notasPrevias) {
      filas.push([nota.hoja, nota.fila, nota.texto]);
    }
  }

  return { filas, merges };
}

function crearHojaResumen(datos) {
  const { filas, merges } = construirContenidoResumen(datos);
  const hoja = XLSX.utils.aoa_to_sheet(filas);
  hoja['!cols'] = [
    { wch: 40 }, { wch: 20 }, { wch: 14 }, { wch: 24 },
    { wch: 24 }, { wch: 26 }, { wch: 24 }
  ];
  hoja['!merges'] = merges;

  return hoja;
}

function normalizarNombreHoja(nombre) {
  return normalizarTexto(nombre) || 'hoja';
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

  while (usados.has(normalizarNombreHoja(nombre))) {
    const sufijo = ` (${contador})`;
    nombre = `${base.slice(0, 31 - sufijo.length)}${sufijo}`;
    contador += 1;
  }

  usados.add(normalizarNombreHoja(nombre));
  return nombre;
}

function agruparPorTipo(registros) {
  const grupos = new Map();

  for (const registro of registros) {
    const label = aTexto(registro.tipo) || 'Sin clasificar';
    const clave = normalizarTexto(label) || 'sin clasificar';

    if (!grupos.has(clave)) {
      grupos.set(clave, { label, registros: [] });
    }

    grupos.get(clave).registros.push(registro);
  }

  return [...grupos.values()].sort((a, b) => {
    if (b.registros.length !== a.registros.length) return b.registros.length - a.registros.length;
    return colador.compare(a.label, b.label);
  });
}

function construirLibroSalida(datos) {
  const libro = XLSX.utils.book_new();
  const nombresUsados = new Set([
    normalizarNombreHoja('Resumen'),
    normalizarNombreHoja('Inventario completo'),
    normalizarNombreHoja('Revisión')
  ]);

  XLSX.utils.book_append_sheet(libro, crearHojaResumen(datos), 'Resumen');
  XLSX.utils.book_append_sheet(
    libro,
    crearHojaDatos(datos.columnas, datos.registros),
    'Inventario completo'
  );

  if (datos.dimensionesGlobales.tipo) {
    for (const grupo of agruparPorTipo(datos.registros)) {
      const nombre = nombreHojaSeguro(grupo.label, nombresUsados);
      XLSX.utils.book_append_sheet(
        libro,
        crearHojaDatos(datos.columnas, grupo.registros),
        nombre
      );
    }
  }

  const registrosRevision = datos.registros.filter((registro) => registro.revision.length > 0);
  if (registrosRevision.length) {
    XLSX.utils.book_append_sheet(
      libro,
      crearHojaDatos(datos.columnas, registrosRevision),
      'Revisión'
    );
  }

  return XLSX.write(libro, {
    type: 'buffer',
    bookType: 'xlsx',
    compression: true,
    cellDates: true
  });
}

function aplicarEstiloEncabezado(fila) {
  fila.eachCell({ includeEmpty: true }, (celda) => {
    celda.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    celda.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1D1E24' }
    };
    celda.alignment = { vertical: 'middle' };
  });
}

function configurarImpresion(hoja, columnas, filas) {
  if (!columnas || !filas) return;

  hoja.pageSetup = {
    paperSize: 9,
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.2,
      footer: 0.2
    }
  };
  hoja.pageSetup.printArea = `A1:${hoja.getColumn(columnas).letter}${filas}`;
}

function crearHojaDatosExcelJS(libro, nombreHoja, columnas, registros) {
  const encabezados = [
    ...columnas.map((columna) => columna.label),
    ...COLUMNAS_DERIVADAS
  ];

  const filas = [encabezados];

  for (const registro of registros) {
    filas.push([
      ...columnas.map((columna) => registro.valores[columna.key] ?? ''),
      registro.hoja,
      registro.filaOriginal,
      registro.tipo,
      registro.sucursal,
      registro.procedencia,
      registro.revision.join('; ')
    ]);
  }

  const hoja = libro.addWorksheet(nombreHoja, {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  filas.forEach((fila, indice) => {
    const filaHoja = hoja.addRow(fila);

    if (indice === 0) {
      aplicarEstiloEncabezado(filaHoja);
      return;
    }

    filaHoja.eachCell({ includeEmpty: true }, (celda) => {
      const valor = celda.value;
      if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
        const tieneHora = valor.getHours() || valor.getMinutes();
        celda.numFmt = tieneHora ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy';
      }
    });
  });

  const anchos = calcularAnchos(encabezados, filas);
  anchos.forEach((ancho, indice) => {
    hoja.getColumn(indice + 1).width = Math.min(Math.max(8, ancho.wch), 45);
  });

  if (filas.length > 1) {
    hoja.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: filas.length, column: encabezados.length }
    };
  }

  configurarImpresion(hoja, encabezados.length, filas.length);

  return hoja;
}

function crearHojaResumenExcelJS(libro, nombreHoja, datos) {
  const { filas, merges } = construirContenidoResumen(datos);
  const hoja = libro.addWorksheet(nombreHoja);

  filas.forEach((fila) => hoja.addRow(fila));

  for (const merge of merges) {
    hoja.mergeCells(merge.s.r + 1, merge.s.c + 1, merge.e.r + 1, merge.e.c + 1);
  }

  const anchos = [40, 20, 14, 24, 24, 26, 24];
  anchos.forEach((ancho, indice) => {
    hoja.getColumn(indice + 1).width = ancho;
  });

  hoja.getRow(1).eachCell({ includeEmpty: true }, (celda) => {
    celda.font = { bold: true, size: 14, color: { argb: 'FFC3002F' } };
  });

  configurarImpresion(hoja, anchos.length, filas.length);

  return hoja;
}

async function construirLibroSalidaFormateado(datos, bufferBase) {
  let libro = new ExcelJS.Workbook();

  if (bufferBase && bufferBase.length) {
    try {
      await libro.xlsx.load(bufferBase);
    } catch (error) {
      libro = new ExcelJS.Workbook();
    }
  }

  const nombresUsados = new Set(
    libro.worksheets.map((hoja) => normalizarNombreHoja(hoja.name))
  );

  const nombreResumen = nombreHojaSeguro('Resumen', nombresUsados);
  crearHojaResumenExcelJS(libro, nombreResumen, datos);

  const nombreCompleto = nombreHojaSeguro('Inventario completo', nombresUsados);
  crearHojaDatosExcelJS(libro, nombreCompleto, datos.columnas, datos.registros);

  if (datos.dimensionesGlobales.tipo) {
    for (const grupo of agruparPorTipo(datos.registros)) {
      const nombre = nombreHojaSeguro(grupo.label, nombresUsados);
      crearHojaDatosExcelJS(libro, nombre, datos.columnas, grupo.registros);
    }
  }

  const registrosRevision = datos.registros.filter((registro) => registro.revision.length > 0);
  if (registrosRevision.length) {
    const nombre = nombreHojaSeguro('Revisión', nombresUsados);
    crearHojaDatosExcelJS(libro, nombre, datos.columnas, registrosRevision);
  }

  const resultado = await libro.xlsx.writeBuffer();
  return Buffer.from(resultado);
}

module.exports = {
  COLUMNAS_DERIVADAS,
  contarValores,
  contarUnicos,
  crearMapaUnion,
  aplicarUnion,
  compararRegistros,
  crearFilasPrevia,
  construirLibroSalida,
  construirLibroSalidaFormateado
};
