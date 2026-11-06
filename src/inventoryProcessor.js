const XLSX = require('xlsx');
const {
  aTexto,
  esVacio,
  detectarFilaEncabezado,
  calcularTotalColumnas,
  construirEncabezados
} = require('./textUtils');
const {
  CATEGORIAS,
  detectarDimensiones,
  crearRevision,
  clasificarCategoria
} = require('./classifier');
const {
  COLUMNAS_DERIVADAS,
  contarValores,
  contarUnicos,
  crearMapaUnion,
  aplicarUnion,
  compararRegistros,
  crearFilasPrevia,
  construirLibroSalidaFormateado
} = require('./workbookBuilder');

const MAX_FILAS_PREVIA = 200;

class ErrorInventario extends Error {
  constructor(mensaje, estado = 400) {
    super(mensaje);
    this.name = 'ErrorInventario';
    this.estado = estado;
  }
}

function crearHojaVacia(nombreHoja) {
  return {
    hoja: nombreHoja,
    vacia: true,
    encabezadoFila: null,
    encabezados: [],
    columnasGeneradas: [],
    columnasRepetidas: [],
    registros: [],
    dimensiones: {
      tipo: null,
      sucursal: null,
      procedencia: null,
      notas: [],
      conflictos: []
    },
    advertencias: [`La hoja «${nombreHoja}» está vacía.`],
    notasPrevias: []
  };
}

function propagarCeldasCombinadas(filas, filasDatos, indiceEncabezado, desplazamientoFilas, desplazamientoColumnas, merges) {
  const copia = filasDatos.map((fila) => (Array.isArray(fila) ? fila.slice() : []));

  if (!Array.isArray(merges) || !merges.length || indiceEncabezado < 0) {
    return { filas: copia, aplicadas: 0 };
  }

  let aplicadas = 0;

  for (const merge of merges) {
    if (merge.e.c !== merge.s.c || merge.e.r <= merge.s.r) continue;

    const filaRelativa = merge.s.r - desplazamientoFilas;
    const columnaRelativa = merge.s.c - desplazamientoColumnas;

    if (filaRelativa <= indiceEncabezado || filaRelativa < 0 || columnaRelativa < 0) continue;

    const valor = filas[filaRelativa] ? filas[filaRelativa][columnaRelativa] : '';
    if (esVacio(valor)) continue;

    const inicio = filaRelativa - (indiceEncabezado + 1);
    const fin = merge.e.r - desplazamientoFilas - (indiceEncabezado + 1);

    for (let i = inicio; i <= fin; i += 1) {
      if (!copia[i]) continue;
      if (esVacio(copia[i][columnaRelativa])) {
        copia[i][columnaRelativa] = valor;
        aplicadas += 1;
      }
    }
  }

  return { filas: copia, aplicadas };
}

function procesarHoja(filas, nombreHoja, desplazamientoFilas, desplazamientoColumnas, merges) {
  const filasSeguras = Array.isArray(filas) ? filas : [];
  const hojaVacia = filasSeguras.length === 0
    || filasSeguras.every((fila) => !Array.isArray(fila) || fila.every(esVacio));

  if (hojaVacia) return crearHojaVacia(nombreHoja);

  const indiceEncabezado = detectarFilaEncabezado(filasSeguras);
  const filaEncabezado = indiceEncabezado >= 0 ? filasSeguras[indiceEncabezado] : [];
  const totalColumnas = calcularTotalColumnas(filasSeguras, filaEncabezado);
  const { encabezados, generadas, repetidas } = construirEncabezados(filaEncabezado, totalColumnas);
  const filasDatos = indiceEncabezado >= 0
    ? filasSeguras.slice(indiceEncabezado + 1)
    : filasSeguras;
  const clasificacion = propagarCeldasCombinadas(
    filasSeguras,
    filasDatos,
    indiceEncabezado,
    desplazamientoFilas,
    desplazamientoColumnas,
    merges
  );
  const filasClasificacion = clasificacion.filas;
  const advertencias = [];

  if (indiceEncabezado < 0) {
    advertencias.push(
      `La hoja «${nombreHoja}» no tiene un encabezado claro. Se generaron columnas genéricas y se conservaron todas las filas.`
    );
  }

  if (generadas.length) {
    advertencias.push(
      `La hoja «${nombreHoja}» tiene ${generadas.length} columna(s) sin encabezado; se nombraron como «Columna N» para poder organizarlas.`
    );
  }

  if (repetidas.length) {
    advertencias.push(
      `La hoja «${nombreHoja}» tiene encabezados repetidos (${[...new Set(repetidas)].join(', ')}); se agregó un número para distinguirlos.`
    );
  }

  if (clasificacion.aplicadas > 0) {
    advertencias.push(
      `En la hoja «${nombreHoja}» se propagaron ${clasificacion.aplicadas} celda(s) combinada(s) para clasificar los registros. Los datos originales se conservan tal como estaban.`
    );
  }

  const dimensiones = detectarDimensiones(encabezados, filasClasificacion);
  const registros = [];

  filasDatos.forEach((fila, indice) => {
    const filaSegura = Array.isArray(fila) ? fila : [];
    const filaClasificacion = Array.isArray(filasClasificacion[indice]) ? filasClasificacion[indice] : filaSegura;
    const cantidadDatos = filaSegura.filter((celda) => !esVacio(celda)).length;
    if (cantidadDatos === 0) return;

    const valoresOriginales = encabezados.map((encabezado) => filaSegura[encabezado.index] ?? '');
    const tipo = dimensiones.tipo ? aTexto(filaClasificacion[dimensiones.tipo.index]) : '';
    const sucursal = dimensiones.sucursal ? aTexto(filaClasificacion[dimensiones.sucursal.index]) : '';
    const procedencia = dimensiones.procedencia ? aTexto(filaClasificacion[dimensiones.procedencia.index]) : '';
    const categoria = clasificarCategoria(tipo, filaClasificacion, nombreHoja);
    const revision = crearRevision(dimensiones, tipo, sucursal, procedencia, filaSegura, cantidadDatos);

    if (categoria === 'Sin clasificar' && !revision.includes('Sin clasificar')) {
      revision.push('Sin clasificar');
    }

    const filaOriginal = indiceEncabezado >= 0
      ? desplazamientoFilas + indiceEncabezado + indice + 2
      : desplazamientoFilas + indice + 1;

    registros.push({
      hoja: nombreHoja,
      filaOriginal,
      valoresOriginales,
      tipo,
      sucursal,
      procedencia,
      categoria,
      revision
    });
  });

  const notasPrevias = [];

  if (indiceEncabezado > 0) {
    filasSeguras.slice(0, indiceEncabezado).forEach((fila, indice) => {
      if (!Array.isArray(fila)) return;
      const valores = fila.map(aTexto).filter(Boolean);
      if (!valores.length) return;

      notasPrevias.push({
        hoja: nombreHoja,
        fila: desplazamientoFilas + indice + 1,
        texto: valores.join(' | ')
      });
    });
  }

  if (!registros.length) {
    advertencias.push(`La hoja «${nombreHoja}» solo contiene encabezados o filas vacías.`);
  }

  return {
    hoja: nombreHoja,
    vacia: false,
    encabezadoFila: indiceEncabezado >= 0 ? desplazamientoFilas + indiceEncabezado + 1 : null,
    encabezados,
    columnasGeneradas: generadas,
    columnasRepetidas: repetidas,
    registros,
    dimensiones,
    advertencias,
    notasPrevias
  };
}

function crearNombreSalida(nombreArchivo) {
  const base = String(nombreArchivo || 'inventario')
    .replace(/\.(xlsx|xls)$/i, '')
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80) || 'inventario';

  return `${base}_organizado.xlsx`;
}

function dimensionDetectada(hoja, dimension) {
  return Boolean(hoja.dimensiones && hoja.dimensiones[dimension]);
}

async function procesarInventario(buffer, nombreArchivo) {
  if (!buffer || !buffer.length) {
    throw new ErrorInventario('El archivo está vacío.');
  }

  let libro;

  try {
    libro = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  } catch (error) {
    throw new ErrorInventario('No se pudo leer el archivo. Verifica que sea un Excel válido (.xlsx o .xls).');
  }

  const nombresHojas = libro.SheetNames || [];

  if (!nombresHojas.length) {
    throw new ErrorInventario('El archivo no contiene hojas de cálculo.');
  }

  const hojas = [];

  for (const nombreHoja of nombresHojas) {
    const worksheet = libro.Sheets[nombreHoja];

    if (!worksheet || !worksheet['!ref']) {
      hojas.push(crearHojaVacia(nombreHoja));
      continue;
    }

    try {
      const rango = XLSX.utils.decode_range(worksheet['!ref']);
      const filas = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: true,
        blankrows: true
      });

      hojas.push(procesarHoja(filas, nombreHoja, rango.s.r, rango.s.c, worksheet['!merges']));
    } catch (error) {
      const hoja = crearHojaVacia(nombreHoja);
      hoja.advertencias = [`No se pudo procesar la hoja «${nombreHoja}».`];
      hojas.push(hoja);
    }
  }

  const mapas = crearMapaUnion(hojas);
  const registrosBase = aplicarUnion(hojas, mapas.mapasPorHoja);

  if (!registrosBase.length) {
    throw new ErrorInventario('No se encontraron filas con datos en el archivo.');
  }

  const dimensionesGlobales = {
    tipo: hojas.some((hoja) => dimensionDetectada(hoja, 'tipo')),
    sucursal: hojas.some((hoja) => dimensionDetectada(hoja, 'sucursal')),
    procedencia: hojas.some((hoja) => dimensionDetectada(hoja, 'procedencia'))
  };

  const clavesOrden = ['categoria', 'tipo', 'sucursal', 'procedencia']
    .filter((dimension) => dimension === 'categoria' || dimensionesGlobales[dimension]);

  if (!clavesOrden.length) clavesOrden.push('hoja');

  const registros = [...registrosBase].sort((a, b) => compararRegistros(a, b, clavesOrden));

  const porTipo = contarValores(registros, 'tipo');
  const porSucursal = contarValores(registros, 'sucursal');
  const porProcedencia = contarValores(registros, 'procedencia');
  const conteoCategorias = new Map(CATEGORIAS.map((label) => [label, 0]));
  let sinClasificar = 0;

  for (const registro of registros) {
    if (conteoCategorias.has(registro.categoria)) {
      conteoCategorias.set(registro.categoria, conteoCategorias.get(registro.categoria) + 1);
    } else {
      sinClasificar += 1;
    }
  }

  const porCategoria = [...conteoCategorias].map(([label, count]) => ({ label, count }));
  if (sinClasificar > 0) {
    porCategoria.push({ label: 'Sin clasificar', count: sinClasificar });
  }

  const paraRevision = registros.filter((registro) => registro.revision.length > 0).length;
  const hojasConDatos = hojas.filter((hoja) => hoja.registros.length > 0).length;
  const hojasVacias = hojas.filter((hoja) => hoja.vacia || hoja.registros.length === 0).length;

  const resumen = {
    totalRegistros: registros.length,
    hojasConDatos,
    hojasVacias,
    paraRevision,
    tiposUnicos: contarUnicos(registros, 'tipo'),
    sucursalesUnicas: contarUnicos(registros, 'sucursal'),
    procedenciasUnicas: contarUnicos(registros, 'procedencia'),
    porTipo,
    porSucursal,
    porProcedencia,
    porCategoria
  };

  const advertencias = [];
  const notasInterpretacion = [];
  const notasPrevias = [];

  for (const hoja of hojas) {
    advertencias.push(...hoja.advertencias);
    notasInterpretacion.push(...hoja.dimensiones.notas);
    advertencias.push(...hoja.dimensiones.conflictos);
    notasPrevias.push(...hoja.notasPrevias);
  }

  const dimensionesNoDetectadas = ['tipo', 'sucursal', 'procedencia']
    .filter((dimension) => !dimensionesGlobales[dimension]);

  if (dimensionesNoDetectadas.length === 3) {
    advertencias.push(
      'No se detectaron columnas de tipo, sucursal ni procedencia. El inventario se conservó completo y se marcó para revisión.'
    );
  } else {
    for (const dimension of dimensionesNoDetectadas) {
      advertencias.push(
        `No se detectó una columna de ${dimension} en ninguna hoja. Los registros se conservan y se marcan para revisión.`
      );
    }
  }

  if (paraRevision > 0) {
    advertencias.push(
      `${paraRevision} registro(s) requieren revisión. Revisa la hoja «Revisión» del Excel generado.`
    );
  }

  const advertenciasUnicas = [...new Set(advertencias)];
  const nombreSalida = crearNombreSalida(nombreArchivo);
  const fechaGeneracion = new Date().toLocaleString('es-MX', {
    dateStyle: 'long',
    timeStyle: 'short'
  });

  const columnasPrevia = [...mapas.columnas, ...COLUMNAS_DERIVADAS.map((label) => ({ label }))];
  const filasPrevia = crearFilasPrevia(registros, mapas.columnas, MAX_FILAS_PREVIA);

  const bufferSalida = await construirLibroSalidaFormateado({
    columnas: mapas.columnas,
    registros,
    resumen,
    hojas,
    advertencias: advertenciasUnicas,
    notasInterpretacion,
    notasPrevias,
    nombreArchivo,
    nombreSalida,
    fechaGeneracion,
    dimensionesGlobales
  });

  const preview = {
    nombreArchivo,
    nombreSalida,
    fechaGeneracion,
    columnas: columnasPrevia.map((columna) => columna.label),
    filas: filasPrevia,
    totalRegistros: registros.length,
    filasMostradas: filasPrevia.length,
    inicioDerivadas: mapas.columnas.length,
    indiceCategoria: columnasPrevia.findIndex((columna) => columna.label === 'Categoría'),
    indiceRevision: columnasPrevia.length - 1,
    pesoSalida: bufferSalida.length,
    resumen: {
      ...resumen,
      dimensiones: dimensionesGlobales
    },
    hojas: hojas.map((hoja) => ({
      nombre: hoja.hoja,
      vacia: hoja.vacia,
      registros: hoja.registros.length,
      encabezadoFila: hoja.encabezadoFila,
      tipo: hoja.dimensiones.tipo ? hoja.dimensiones.tipo.label : null,
      sucursal: hoja.dimensiones.sucursal ? hoja.dimensiones.sucursal.label : null,
      procedencia: hoja.dimensiones.procedencia ? hoja.dimensiones.procedencia.label : null,
      columnasGeneradas: hoja.columnasGeneradas.length
    })),
    advertencias: advertenciasUnicas,
    notasInterpretacion,
    notasPrevias
  };

  return {
    buffer: bufferSalida,
    nombreSalida,
    preview
  };
}

module.exports = {
  procesarInventario,
  ErrorInventario
};
