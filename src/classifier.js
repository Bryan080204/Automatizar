const {
  PISTAS_NACIONALIDAD,
  normalizarTexto,
  contieneFrase,
  aTexto,
  contarPistasEncabezado
} = require('./textUtils');

const CATEGORIAS = ['Nacionales', 'Importados', 'Internacionales', 'NP300'];

const REGLAS_DIMENSION = {
  tipo: [
    ['tipo', 100],
    ['clasificacion', 95],
    ['categoria', 90],
    ['nacional', 85],
    ['importado', 85],
    ['internacional', 85],
    ['extranjero', 85],
    ['nacionalidad', 80],
    ['condicion', 55],
    ['origen', 35]
  ],
  sucursal: [
    ['sucursal', 100],
    ['agencia', 95],
    ['concesionario', 90],
    ['distribuidor', 85],
    ['dealer', 80],
    ['punto de venta', 80],
    ['sala de venta', 80],
    ['tienda', 70]
  ],
  procedencia: [
    ['procedencia', 100],
    ['ciudad', 95],
    ['localidad', 90],
    ['municipio', 88],
    ['poblacion', 85],
    ['plaza', 70],
    ['region', 60],
    ['entidad', 55],
    ['pais', 50],
    ['zona', 50],
    ['lugar', 45],
    ['origen', 35]
  ]
};

function prioridadDimension(dimension) {
  if (dimension === 'tipo') return 0;
  if (dimension === 'sucursal') return 1;
  return 2;
}

function puntuarDimension(textoNormalizado, dimension) {
  const reglas = REGLAS_DIMENSION[dimension] || [];
  let puntaje = 0;
  let posicion = Infinity;

  for (const [frase, puntos] of reglas) {
    if (!contieneFrase(textoNormalizado, frase)) continue;
    posicion = Math.min(posicion, textoNormalizado.indexOf(frase));
    puntaje = Math.max(puntaje, puntos);
  }

  return { puntaje, posicion };
}

function esEncabezadoSoloOrigen(textoNormalizado) {
  if (!contieneFrase(textoNormalizado, 'origen')) return false;

  const palabrasEspecificas = [
    'tipo', 'clasificacion', 'categoria', 'nacional', 'importado',
    'internacional', 'extranjero', 'sucursal', 'agencia', 'concesionario',
    'distribuidor', 'procedencia', 'ciudad', 'localidad', 'municipio',
    'poblacion', 'pais'
  ];

  return !palabrasEspecificas.some((palabra) => contieneFrase(textoNormalizado, palabra));
}

function valoresParecenNacionalidad(filasDatos, indiceColumna) {
  const valores = [];

  for (const fila of filasDatos) {
    const valor = normalizarTexto(Array.isArray(fila) ? fila[indiceColumna] : '');
    if (valor) valores.push(valor);
    if (valores.length >= 40) break;
  }

  if (!valores.length) return false;

  let coincidencias = 0;
  for (const valor of valores) {
    const tokens = valor.split(' ').filter(Boolean);
    if (PISTAS_NACIONALIDAD.has(valor) || tokens.some((token) => PISTAS_NACIONALIDAD.has(token))) {
      coincidencias += 1;
    }
  }

  return coincidencias / valores.length >= 0.6;
}

function detectarDimensiones(encabezados, filasDatos) {
  const candidatos = { tipo: [], sucursal: [], procedencia: [] };
  const notas = [];
  const conflictos = [];

  for (const encabezado of encabezados) {
    if (encabezado.generada) continue;

    const texto = normalizarTexto(encabezado.labelBase);
    if (!texto) continue;

    let seleccion;

    if (esEncabezadoSoloOrigen(texto)) {
      const nacionalidad = valoresParecenNacionalidad(filasDatos, encabezado.index);
      seleccion = nacionalidad
        ? { dimension: 'tipo', puntaje: 75, posicion: 0 }
        : { dimension: 'procedencia', puntaje: 65, posicion: 0 };
      notas.push(
        `La columna «${encabezado.label}» se interpretó como ${seleccion.dimension === 'tipo' ? 'tipo de vehículo' : 'procedencia'} según sus valores.`
      );
    } else {
      const evaluaciones = Object.keys(REGLAS_DIMENSION).map((dimension) => ({
        dimension,
        ...puntuarDimension(texto, dimension)
      }));

      evaluaciones.sort((a, b) => {
        if (a.posicion !== b.posicion) return a.posicion - b.posicion;
        if (a.puntaje !== b.puntaje) return b.puntaje - a.puntaje;
        return prioridadDimension(a.dimension) - prioridadDimension(b.dimension);
      });

      seleccion = evaluaciones[0];
    }

    if (!seleccion || seleccion.puntaje <= 0) continue;
    candidatos[seleccion.dimension].push({ encabezado, ...seleccion });
  }

  const resultado = {
    tipo: null,
    sucursal: null,
    procedencia: null,
    notas,
    conflictos
  };

  for (const dimension of ['tipo', 'sucursal', 'procedencia']) {
    const lista = candidatos[dimension].sort((a, b) => {
      if (b.puntaje !== a.puntaje) return b.puntaje - a.puntaje;
      if (a.posicion !== b.posicion) return a.posicion - b.posicion;
      return a.encabezado.index - b.encabezado.index;
    });

    if (!lista.length) continue;
    resultado[dimension] = lista[0].encabezado;

    if (lista.length > 1) {
      const alternativas = lista.slice(1).map((item) => `«${item.encabezado.label}»`).join(', ');
      conflictos.push(
        `Varias columnas podrían ser ${dimension}: se usó «${lista[0].encabezado.label}» y se dejaron como datos originales ${alternativas}.`
      );
    }
  }

  return resultado;
}

function crearRevision(dimensiones, tipo, sucursal, procedencia, fila, cantidadDatos) {
  const revision = [];
  const sinTipo = !dimensiones.tipo || !tipo;
  const sinSucursal = !dimensiones.sucursal || !sucursal;
  const sinProcedencia = !dimensiones.procedencia || !procedencia;

  if (sinTipo && sinSucursal && sinProcedencia) {
    revision.push('Sin clasificación');
  } else {
    if (sinTipo) revision.push('Sin tipo');
    if (sinSucursal) revision.push('Sin sucursal');
    if (sinProcedencia) revision.push('Sin procedencia');
  }

  if (cantidadDatos <= 1) {
    revision.push('Fila con un solo dato (posible título o registro incompleto)');
  }

  const pistasFila = fila.reduce((total, celda) => total + contarPistasEncabezado(celda), 0);
  if (pistasFila >= 4 && cantidadDatos >= 3) {
    revision.push('Posible encabezado repetido');
  }

  return revision;
}

function pareceNP300(valor) {
  const texto = normalizarTexto(valor);
  if (!texto) return false;

  return /(^|\s)np\s*3\s*[0o]\s*[0o](\s|$)/.test(texto);
}

function categoriaPorTexto(valor) {
  const texto = normalizarTexto(valor);
  if (!texto) return null;

  if (texto.includes('internacional')) return 'Internacionales';
  if (
    texto.includes('importado')
    || texto.includes('importacion')
    || texto.includes('extranjero')
  ) {
    return 'Importados';
  }
  if (
    texto.includes('nacional')
    || texto.includes('nacionalizado')
    || texto.includes('local')
  ) {
    return 'Nacionales';
  }

  return null;
}

function clasificarCategoria(tipo, valores, pista) {
  const textos = [tipo, pista, ...(Array.isArray(valores) ? valores : [])];

  if (textos.some(pareceNP300)) return 'NP300';

  return categoriaPorTexto(tipo) || categoriaPorTexto(pista) || 'Sin clasificar';
}

module.exports = {
  CATEGORIAS,
  detectarDimensiones,
  crearRevision,
  clasificarCategoria
};
