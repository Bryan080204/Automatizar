const XLSX = require('xlsx');
const { aTexto, normalizarTexto } = require('./textUtils');
const {
  CATEGORIAS,
  CATEGORIA_DETALLE,
  clasificarVehiculo,
  mapearUbicacion,
  detectarSeccion
} = require('./classifier');
const {
  COLUMNAS_SALIDA,
  construirLibroSalida,
  contarPor
} = require('./workbookBuilder');

const MAX_FILAS_PREVIA = 200;

const ALIAS_COLUMNAS = {
  sucursal: ['sucursal'],
  mod: ['mod', 'modelo', 'ano', 'año'],
  descripcion: ['descripcion local', 'descripcion', 'descripcion del vehiculo', 'version'],
  color: ['colores', 'color'],
  serie: ['serie', 'vin', 'numero de serie', 'no serie'],
  dealer: ['dealer', 'distribuidor', 'concesionario'],
  dias: ['cod', 'dias', 'antiguedad', 'dias de inventario', 'dias en piso'],
  fecha: ['ant', 'fecha', 'fecha de ingreso', 'fecha de llegada'],
  observaciones: ['sistema de apartado', 'observaciones', 'observacion', 'notas', 'comentarios']
};

class ErrorInventario extends Error {
  constructor(mensaje, estado = 400) {
    super(mensaje);
    this.name = 'ErrorInventario';
    this.estado = estado;
  }
}

function detectarEncabezadoNissan(filas) {
  const limite = Math.min(filas.length, 40);
  let mejor = { indice: -1, puntaje: 0 };

  for (let indice = 0; indice < limite; indice += 1) {
    const fila = Array.isArray(filas[indice]) ? filas[indice] : [];
    const celdas = fila.map(normalizarTexto);
    let puntaje = 0;

    if (celdas.includes('sucursal')) puntaje += 2;
    if (celdas.some((celda) => celda === 'descripcion local' || celda === 'descripcion')) puntaje += 3;
    if (celdas.includes('serie') || celdas.includes('vin')) puntaje += 2;
    if (celdas.includes('dealer')) puntaje += 2;
    if (celdas.includes('mod')) puntaje += 2;
    if (celdas.includes('ant')) puntaje += 2;
    if (celdas.includes('cod')) puntaje += 2;
    if (celdas.includes('colores') || celdas.includes('color')) puntaje += 1;

    if (puntaje > mejor.puntaje) {
      mejor = { indice, puntaje };
    }
  }

  return mejor.puntaje >= 8 ? mejor.indice : -1;
}

function indexarColumnas(filaEncabezado) {
  const celdas = (Array.isArray(filaEncabezado) ? filaEncabezado : []).map(normalizarTexto);

  const buscar = (aliases) => {
    for (const alias of aliases) {
      const indice = celdas.indexOf(normalizarTexto(alias));
      if (indice >= 0) return indice;
    }
    return -1;
  };

  const indices = {
    sucursal: buscar(ALIAS_COLUMNAS.sucursal),
    mod: buscar(ALIAS_COLUMNAS.mod),
    descripcion: buscar(ALIAS_COLUMNAS.descripcion),
    color: buscar(ALIAS_COLUMNAS.color),
    serie: buscar(ALIAS_COLUMNAS.serie),
    dealer: buscar(ALIAS_COLUMNAS.dealer),
    dias: buscar(ALIAS_COLUMNAS.dias),
    fecha: buscar(ALIAS_COLUMNAS.fecha),
    observaciones: buscar(ALIAS_COLUMNAS.observaciones)
  };

  const etiquetas = {};
  for (const [clave, indice] of Object.entries(indices)) {
    etiquetas[clave] = indice >= 0 ? aTexto(filaEncabezado[indice]) : null;
  }

  return { indices, etiquetas };
}

function parseFecha(valor) {
  if (valor instanceof Date) {
    return Number.isNaN(valor.getTime()) ? null : valor;
  }

  if (typeof valor === 'number' && Number.isFinite(valor)) {
    const partes = XLSX.SSF.parse_date_code(valor);
    if (!partes || !partes.y) return null;
    return new Date(Date.UTC(partes.y, partes.m - 1, partes.d));
  }

  const texto = aTexto(valor);
  if (!texto) return null;

  let coincidencia = texto.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
  if (coincidencia) {
    let anio = Number(coincidencia[3]);
    if (anio < 100) anio += 2000;
    return new Date(Date.UTC(anio, Number(coincidencia[2]) - 1, Number(coincidencia[1])));
  }

  coincidencia = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (coincidencia) {
    return new Date(Date.UTC(Number(coincidencia[1]), Number(coincidencia[2]) - 1, Number(coincidencia[3])));
  }

  return null;
}

function detectarFechaCorte(filas, indiceEncabezado) {
  const limite = Math.max(0, indiceEncabezado);

  for (let indice = 0; indice <= limite; indice += 1) {
    const fila = Array.isArray(filas[indice]) ? filas[indice] : [];

    for (const celda of fila) {
      if (celda instanceof Date && !Number.isNaN(celda.getTime())) return celda;

      if (typeof celda === 'string' && /\d{1,4}[/-]\d{1,2}[/-]\d{1,4}/.test(celda)) {
        const fecha = parseFecha(celda);
        if (fecha) return fecha;
      }
    }
  }

  return null;
}

function normalizarDia(fecha) {
  return Date.UTC(fecha.getUTCFullYear(), fecha.getUTCMonth(), fecha.getUTCDate());
}

function diasEntre(fecha, corte) {
  if (!fecha || !corte) return 0;
  const diferencia = Math.floor((normalizarDia(corte) - normalizarDia(fecha)) / 86400000);
  return Math.max(0, diferencia);
}

function esFilaTotal(fila) {
  const celdas = (Array.isArray(fila) ? fila : []).slice(0, 4).map(normalizarTexto);

  return celdas.some((celda) => (
    /^(sub[\s-]*total|total|unidades\s+por\s+pagar)(\s|$)/.test(celda)
    || /^total:?$/.test(celda)
  ));
}

function limpiarObservaciones(valor) {
  const texto = aTexto(valor).replace(/\s+/g, ' ').trim();
  if (!texto) return '';

  const digitos = (texto.match(/\d/g) || []).length;
  if (digitos >= 8 && /^\+?[\d\s().-]+$/.test(texto)) return '';

  return texto;
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

function procesarHoja(filas, nombreHoja, desplazamientoFilas) {
  const filasSeguras = Array.isArray(filas) ? filas : [];
  const indiceEncabezado = detectarEncabezadoNissan(filasSeguras);

  if (indiceEncabezado < 0) {
    return {
      valida: false,
      nombre: nombreHoja,
      motivo: 'No se encontró el encabezado de la LISTA DE EXISTENCIAS.',
      registros: [],
      secciones: [],
      incluidos: 0,
      excluidos: 0,
      advertencias: []
    };
  }

  const filaEncabezado = filasSeguras[indiceEncabezado];
  const { indices, etiquetas } = indexarColumnas(filaEncabezado);

  if (indices.descripcion < 0 && indices.serie < 0) {
    return {
      valida: false,
      nombre: nombreHoja,
      motivo: 'La hoja tiene encabezado, pero no se detectaron las columnas de Descripción ni Serie.',
      registros: [],
      secciones: [],
      incluidos: 0,
      excluidos: 0,
      advertencias: []
    };
  }

  const fechaCorte = detectarFechaCorte(filasSeguras, indiceEncabezado);
  const seccionesMap = new Map();
  const registros = [];
  const advertencias = [];
  let sinSerie = 0;
  let sinClasificar = 0;
  let sinDias = 0;

  let seccionActual = 'GENERAL';
  if (!seccionesMap.has(seccionActual)) {
    seccionesMap.set(seccionActual, { label: seccionActual, incluida: true, count: 0 });
  }

  for (let indice = indiceEncabezado + 1; indice < filasSeguras.length; indice += 1) {
    const fila = Array.isArray(filasSeguras[indice]) ? filasSeguras[indice] : [];
    const seccion = detectarSeccion(fila);

    if (seccion) {
      seccionActual = seccion.label;
      if (!seccionesMap.has(seccionActual)) {
        seccionesMap.set(seccionActual, {
          label: seccion.label,
          incluida: seccion.incluida,
          count: 0
        });
      }
      continue;
    }

    if (esFilaTotal(fila)) continue;

    const descripcion = indices.descripcion >= 0 ? aTexto(fila[indices.descripcion]) : '';
    const serie = indices.serie >= 0 ? aTexto(fila[indices.serie]) : '';
    const sucursal = indices.sucursal >= 0 ? aTexto(fila[indices.sucursal]) : '';
    const dealer = indices.dealer >= 0 ? aTexto(fila[indices.dealer]) : '';
    const color = indices.color >= 0 ? aTexto(fila[indices.color]) : '';
    const mod = indices.mod >= 0 ? aTexto(fila[indices.mod]) : '';
    const observacionesOriginales = indices.observaciones >= 0 ? aTexto(fila[indices.observaciones]) : '';

    if (!descripcion && !serie) continue;

    const descripcionNormalizada = normalizarTexto(descripcion);
    if (
      descripcionNormalizada === 'descripcion local'
      || descripcionNormalizada === 'sucursal'
      || descripcionNormalizada === 'mod'
    ) {
      continue;
    }

    const seccionInfo = seccionesMap.get(seccionActual) || {
      label: seccionActual,
      incluida: true,
      count: 0
    };
    seccionInfo.count += 1;
    seccionesMap.set(seccionActual, seccionInfo);

    if (!seccionInfo.incluida) continue;

    const diasCrudos = indices.dias >= 0 ? Number(fila[indices.dias]) : NaN;
    const fechaAnt = indices.fecha >= 0 ? parseFecha(fila[indices.fecha]) : null;
    let dias = Number.isFinite(diasCrudos) && diasCrudos >= 0 ? Math.round(diasCrudos) : null;

    if (dias === null) {
      if (fechaAnt && fechaCorte) {
        dias = diasEntre(fechaAnt, fechaCorte);
      } else {
        dias = 0;
        sinDias += 1;
      }
    }

    if (!serie) sinSerie += 1;

    const clasificacion = clasificarVehiculo(descripcion);
    if (clasificacion.categoria === 'Sin clasificar') sinClasificar += 1;

    registros.push({
      hoja: nombreHoja,
      filaOriginal: desplazamientoFilas + indice + 1,
      seccion: seccionActual,
      ubicacion: mapearUbicacion(dealer, sucursal),
      sucursal,
      mod,
      descripcion,
      color,
      serie,
      dealer,
      dias,
      clasificacion: clasificacion.clasificacion,
      categoria: clasificacion.categoria,
      modelo: clasificacion.modelo,
      observaciones: limpiarObservaciones(observacionesOriginales)
    });
  }

  const secciones = [...seccionesMap.values()].filter((seccion) => seccion.count > 0);
  const incluidos = registros.length;
  const excluidos = secciones
    .filter((seccion) => !seccion.incluida)
    .reduce((total, seccion) => total + seccion.count, 0);

  if (!fechaCorte && registros.length) {
    advertencias.push(
      `No se detectó una fecha de corte en la hoja «${nombreHoja}»; los días faltantes se calcularon con la fecha actual.`
    );
  }

  if (sinSerie > 0) {
    advertencias.push(
      `La hoja «${nombreHoja}» tiene ${sinSerie} registro(s) sin Serie. Se conservaron para no perder información.`
    );
  }

  if (indices.dealer < 0) {
    advertencias.push(
      `La hoja «${nombreHoja}» no tiene columna Dealer; la ubicación de origen se tomó de Sucursal.`
    );
  }

  if (indices.dias < 0) {
    advertencias.push(
      `La hoja «${nombreHoja}» no tiene columna Cod.; los días se calcularon desde la fecha Ant. y la fecha de corte.`
    );
  }

  return {
    valida: true,
    nombre: nombreHoja,
    encabezadoFila: desplazamientoFilas + indiceEncabezado + 1,
    fechaCorte,
    columnas: etiquetas,
    registros,
    secciones,
    incluidos,
    excluidos,
    sinSerie,
    sinClasificar,
    advertencias
  };
}

function formatearFechaCorte(fecha) {
  if (!fecha) return 'No detectada';
  return aTexto(fecha);
}

function contarPorUbicacion(registros) {
  const categorias = [...CATEGORIAS, 'Sin clasificar'];
  const mapa = new Map();

  for (const registro of registros) {
    if (!mapa.has(registro.ubicacion)) {
      mapa.set(registro.ubicacion, {
        label: registro.ubicacion,
        count: 0,
        categorias: Object.fromEntries(categorias.map((categoria) => [categoria, 0]))
      });
    }

    const item = mapa.get(registro.ubicacion);
    item.count += 1;

    if (Object.prototype.hasOwnProperty.call(item.categorias, registro.categoria)) {
      item.categorias[registro.categoria] += 1;
    } else {
      item.categorias['Sin clasificar'] += 1;
    }
  }

  return [...mapa.values()].sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.label.localeCompare(b.label, 'es', { numeric: true });
  });
}

function construirResumen(registros, hojasValidas, totalExcluidos) {
  const total = registros.length;
  const conteoCategorias = new Map(CATEGORIAS.map((categoria) => [categoria, 0]));

  for (const registro of registros) {
    if (conteoCategorias.has(registro.categoria)) {
      conteoCategorias.set(registro.categoria, conteoCategorias.get(registro.categoria) + 1);
    }
  }

  const sinClasificar = registros.filter((registro) => registro.categoria === 'Sin clasificar').length;
  const porCategoria = [...conteoCategorias].map(([label, count]) => ({ label, count }));
  if (sinClasificar > 0) porCategoria.push({ label: 'Sin clasificar', count: sinClasificar });

  const seccionesMap = new Map();
  for (const hoja of hojasValidas) {
    for (const seccion of hoja.secciones) {
      if (!seccionesMap.has(seccion.label)) {
        seccionesMap.set(seccion.label, {
          label: seccion.label,
          incluida: seccion.incluida,
          count: 0
        });
      }
      seccionesMap.get(seccion.label).count += seccion.count;
    }
  }

  const secciones = [...seccionesMap.values()];
  const dias = registros.map((registro) => registro.dias).filter((valor) => Number.isFinite(valor));
  const fechaCorte = hojasValidas.find((hoja) => hoja.fechaCorte)?.fechaCorte || null;
  const ubicacionesUnicas = new Set(registros.map((registro) => registro.ubicacion)).size;
  const modelosUnicos = new Set(registros.map((registro) => registro.modelo)).size;

  return {
    totalRegistros: total,
    totalExcluidos,
    sinClasificar,
    porCategoria,
    porUbicacion: contarPorUbicacion(registros),
    porModelo: contarPor(registros, 'modelo'),
    porClasificacion: contarPor(registros, 'clasificacion'),
    secciones,
    fechaCorte: formatearFechaCorte(fechaCorte),
    diasMax: dias.length ? Math.max(...dias) : 0,
    diasMin: dias.length ? Math.min(...dias) : 0,
    promedioDias: dias.length ? Math.round(dias.reduce((suma, valor) => suma + valor, 0) / dias.length) : 0,
    ubicacionesUnicas,
    modelosUnicos,
    hojasProcesadas: hojasValidas.length
  };
}

function construirNotas(resumen) {
  const incluidas = resumen.secciones.filter((seccion) => seccion.incluida);
  const excluidas = resumen.secciones.filter((seccion) => !seccion.incluida);

  const notas = [];

  notas.push(
    `Se conservaron las secciones ${incluidas.map((seccion) => `${seccion.label} (${seccion.count})`).join(', ') || 'sin secciones marcadas'}.`
  );

  notas.push(
    excluidas.length
      ? `Se excluyeron las secciones ${excluidas.map((seccion) => `${seccion.label} (${seccion.count})`).join(', ')}: ${resumen.totalExcluidos} registro(s) de secciones excluidas (seminuevos, usados o Tlalixtac).`
      : 'No se encontraron secciones de seminuevos, usados o Tlalixtac en el archivo.'
  );

  notas.push('Los registros se ordenaron de mayor a menor antigüedad con base en la columna Días.');
  notas.push('La ubicación de origen se tomó del código del Dealer: 300 → UNIVERSIDAD, 287 → MADERO, 298 → PUERTO ESCONDIDO, 301 → HUATULCO, 66 → HUAJUAPAN. El código 300 con FLOTILLA/FLOTIL se muestra como FLOTILLA.');
  notas.push('Clasificación: I = Importado, C = Compacto, D = Doméstico / NP300, S = Subcompacto.');
  notas.push('Las Frontier con motor V6, híbrido o e-POWER se clasifican como importadas; las Frontier de 4 cilindros se clasifican como D / NP300.');
  notas.push('El libro resultado incluye una hoja independiente por cada ubicación de origen.');

  return notas;
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
  const hojasValidas = [];
  const registros = [];

  for (const nombreHoja of nombresHojas) {
    const worksheet = libro.Sheets[nombreHoja];

    if (!worksheet || !worksheet['!ref']) {
      hojas.push({
        valida: false,
        nombre: nombreHoja,
        motivo: 'La hoja está vacía.',
        incluidos: 0,
        excluidos: 0,
        secciones: [],
        registros: []
      });
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

      const resultado = procesarHoja(filas, nombreHoja, rango.s.r);
      hojas.push(resultado);

      if (resultado.valida) {
        hojasValidas.push(resultado);
        registros.push(...resultado.registros);
      }
    } catch (error) {
      hojas.push({
        valida: false,
        nombre: nombreHoja,
        motivo: 'No se pudo procesar la hoja.',
        incluidos: 0,
        excluidos: 0,
        secciones: [],
        registros: []
      });
    }
  }

  if (!hojasValidas.length) {
    throw new ErrorInventario(
      'No se encontró una hoja con el formato de LISTA DE EXISTENCIAS. Verifica que el archivo tenga las columnas Sucursal, Descripción, Serie, Dealer, Ant. y Cod.'
    );
  }

  if (!registros.length) {
    const excluidos = hojasValidas.reduce((total, hoja) => total + hoja.excluidos, 0);
    if (excluidos > 0) {
      throw new ErrorInventario(
        'El archivo solo contiene secciones de seminuevos o usados. No hay unidades de APARTADO, DEMOSTRACION o PISO para generar el reporte.'
      );
    }
    throw new ErrorInventario('No se encontraron registros con datos en la LISTA DE EXISTENCIAS.');
  }

  const registrosOrdenados = [...registros].sort((a, b) => {
    if (b.dias !== a.dias) return b.dias - a.dias;
    return a.filaOriginal - b.filaOriginal;
  });

  const totalExcluidos = hojasValidas.reduce((total, hoja) => total + hoja.excluidos, 0);
  const resumenBase = construirResumen(registrosOrdenados, hojasValidas, totalExcluidos);
  const notasInterpretacion = construirNotas(resumenBase);

  const advertencias = [];
  for (const hoja of hojas) {
    if (hoja.valida) {
      advertencias.push(...hoja.advertencias);
    } else if (hoja.motivo) {
      advertencias.push(`La hoja «${hoja.nombre}» se omitió: ${hoja.motivo}`);
    }
  }

  if (resumenBase.sinClasificar > 0) {
    advertencias.push(
      `${resumenBase.sinClasificar} registro(s) no coinciden con el catálogo de modelos y se conservaron en la hoja «SIN CLASIFICAR».`
    );
  }

  if (!hojasValidas.some((hoja) => hoja.fechaCorte)) {
    advertencias.push('No se detectó la fecha de corte; revisa la columna Ant. si necesitas calcular la antigüedad.');
  }

  const advertenciasUnicas = [...new Set(advertencias)];
  const nombreSalida = crearNombreSalida(nombreArchivo);
  const fechaGeneracion = new Date().toLocaleString('es-MX', {
    dateStyle: 'long',
    timeStyle: 'short'
  });

  const resumen = {
    ...resumenBase,
    porCategoria: [
      ...CATEGORIAS.map((categoria) => ({
        label: categoria,
        count: resumenBase.porCategoria.find((item) => item.label === categoria)?.count || 0
      })),
      ...(resumenBase.sinClasificar > 0
        ? [{ label: 'Sin clasificar', count: resumenBase.sinClasificar }]
        : [])
    ]
  };

  const bufferSalida = await construirLibroSalida({
    registros: registrosOrdenados,
    resumen,
    hojas,
    advertencias: advertenciasUnicas,
    notasInterpretacion,
    nombreArchivo,
    nombreSalida,
    fechaGeneracion
  });

  const filasPrevia = registrosOrdenados.slice(0, MAX_FILAS_PREVIA).map((registro) => (
    COLUMNAS_SALIDA.map((columna) => registro[columna.key])
  ));

  const preview = {
    nombreArchivo,
    nombreSalida,
    fechaGeneracion,
    columnas: COLUMNAS_SALIDA.map((columna) => columna.label),
    filas: filasPrevia,
    totalRegistros: registrosOrdenados.length,
    filasMostradas: filasPrevia.length,
    indiceCategoria: COLUMNAS_SALIDA.findIndex((columna) => columna.key === 'categoria'),
    indiceClasificacion: COLUMNAS_SALIDA.findIndex((columna) => columna.key === 'clasificacion'),
    pesoSalida: bufferSalida.length,
    resumen,
    hojas: hojas.map((hoja) => ({
      nombre: hoja.nombre,
      valida: hoja.valida,
      encabezadoFila: hoja.encabezadoFila || null,
      registros: hoja.incluidos || 0,
      excluidos: hoja.excluidos || 0,
      secciones: hoja.secciones || [],
      columnas: hoja.columnas || null,
      motivo: hoja.motivo || ''
    })),
    advertencias: advertenciasUnicas,
    notasInterpretacion
  };

  return {
    buffer: bufferSalida,
    nombreSalida,
    preview
  };
}

module.exports = {
  procesarInventario,
  ErrorInventario,
  CATEGORIA_DETALLE
};
