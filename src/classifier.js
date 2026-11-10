const { normalizarTexto, aTexto } = require('./textUtils');

/**
 * Catálogo de clasificación para la LISTA DE EXISTENCIAS de Nissan.
 *
 * Las letras solicitadas son:
 *   I = Importado
 *   C = Compacto
 *   D = Doméstico / NP300
 *   S = Subcompacto
 */

const CATEGORIAS = ['Importados', 'Compactos', 'NP300', 'Subcompactos'];

const CATEGORIA_DETALLE = {
  Importados: {
    letra: 'I',
    color: 'FFC3002F',
    fondo: 'FFFFE9EF',
    colorSuave: 'FFF7DCE4',
    icono: 'ship'
  },
  Compactos: {
    letra: 'C',
    color: 'FF2563EB',
    fondo: 'FFE8EFFF',
    colorSuave: 'FFDCE6FF',
    icono: 'car'
  },
  NP300: {
    letra: 'D',
    color: 'FFB45309',
    fondo: 'FFFFF1E0',
    colorSuave: 'FFFBE3C5',
    icono: 'truck'
  },
  Subcompactos: {
    letra: 'S',
    color: 'FF0F7B53',
    fondo: 'FFE7F6EE',
    colorSuave: 'FFD2EEDF',
    icono: 'small-car'
  },
  'Sin clasificar': {
    letra: '?',
    color: 'FF687182',
    fondo: 'FFF3F4F7',
    colorSuave: 'FFE5E8ED',
    icono: 'alert'
  }
};

const ORDEN_CATEGORIA = new Map(CATEGORIAS.map((label, indice) => [label, indice]));

const UBICACIONES_POR_CODIGO = Object.freeze({
  300: 'UNIVERSIDAD',
  287: 'MADERO',
  298: 'PUERTO ESCONDIDO',
  301: 'HUATULCO',
  66: 'HUAJUAPAN'
});

const SECCIONES = Object.freeze({
  APARTADO: { label: 'APARTADO', incluida: true },
  DEMOSTRACION: { label: 'DEMOSTRACION', incluida: true },
  PISO: { label: 'PISO', incluida: true },
  GENERAL: { label: 'GENERAL', incluida: true },
  'SEMI_NUEVOS GARANTIZADOS': { label: 'SEMI_NUEVOS GARANTIZADOS', incluida: false },
  SEMINUEVOS: { label: 'SEMINUEVOS', incluida: false },
  TLALIXTAC: { label: 'TLALIXTAC', incluida: false },
  USADOS: { label: 'USADOS', incluida: false }
});

const SECCIONES_INCLUIDAS = new Set(
  Object.values(SECCIONES).filter((seccion) => seccion.incluida).map((seccion) => seccion.label)
);

const ALIAS_SECCION = new Map([
  ['apartado', 'APARTADO'],
  ['demostracion', 'DEMOSTRACION'],
  ['piso', 'PISO'],
  ['general', 'GENERAL'],
  ['semi nuevos garantizados', 'SEMI_NUEVOS GARANTIZADOS'],
  ['semi nuevo garantizado', 'SEMI_NUEVOS GARANTIZADOS'],
  ['seminuevos', 'SEMINUEVOS'],
  ['seminuevo', 'SEMINUEVOS'],
  ['semi nuevos', 'SEMINUEVOS'],
  ['semi nuevo', 'SEMINUEVOS'],
  ['tlalixtac', 'TLALIXTAC'],
  ['usados', 'USADOS'],
  ['usado', 'USADOS']
]);

const MODELOS = [
  { modelo: 'NP300', regex: /(^|\s)np\s*3\s*[0o]\s*[0o](\s|$)/ },
  { modelo: 'X-TRAIL', regex: /(^|\s)x\s*-?\s*trail(\s|$)/ },
  { modelo: 'V-DRIVE', regex: /(^|\s)v\s*-?\s*drive(\s|$)/ },
  { modelo: 'MAGNITE', regex: /(^|\s)magnite(\s|$)/ },
  { modelo: 'FRONTIER', regex: /(^|\s)frontier(\s|$)/ },
  { modelo: 'KICKS', regex: /(^|\s)kicks(\s|$)/ },
  { modelo: 'VERSA', regex: /(^|\s)versa(\s|$)/ },
  { modelo: 'SENTRA', regex: /(^|\s)sentra(\s|$)/ },
  { modelo: 'MARCH', regex: /(^|\s)march(\s|$)/ },
  { modelo: 'KAIT', regex: /(^|\s)kait(\s|$)/ },
  { modelo: 'URVAN', regex: /(^|\s)urvan(\s|$)/ },
  { modelo: 'PATHFINDER', regex: /(^|\s)pathfinder(\s|$)/ },
  { modelo: 'TIIDA', regex: /(^|\s)tiida(\s|$)/ },
  { modelo: 'NOTE', regex: /(^|\s)note(\s|$)/ },
  { modelo: 'MICRA', regex: /(^|\s)micra(\s|$)/ },
  { modelo: 'Z', regex: /(^|\s)z(\s|$)/ }
];

const MODELOS_IMPORTADOS = new Set([
  'MAGNITE',
  'X-TRAIL',
  'KAIT',
  'URVAN',
  'PATHFINDER',
  'Z'
]);

const MODELOS_COMPACTOS = new Set([
  'VERSA',
  'SENTRA',
  'V-DRIVE',
  'TIIDA',
  'NOTE'
]);

const MODELOS_SUBCOMPACTOS = new Set([
  'MARCH',
  'MICRA'
]);

function pareceNP300(valor) {
  const texto = normalizarTexto(valor);
  if (!texto) return false;
  return /(^|\s)np\s*3\s*[0o]\s*[0o](\s|$)/.test(texto);
}

function detectarModelo(valor) {
  const texto = normalizarTexto(valor);
  if (!texto) return '';

  for (const regla of MODELOS) {
    if (regla.regex.test(texto)) return regla.modelo;
  }

  return '';
}

function clasificarVehiculo(valor) {
  const texto = normalizarTexto(valor);
  const modelo = detectarModelo(texto);
  let categoria = 'Sin clasificar';

  if (pareceNP300(texto)) {
    categoria = 'NP300';
  } else if (modelo === 'FRONTIER') {
    // Las Frontier V6, PRO-4X (V6), PRO HÍBRIDO y PLUG-IN son importadas.
    // Las Frontier regulares (4 cilindros) se consideran NP300.
    categoria = /(^|\s)v6(\s|$)|pro\s*hibrido|plug\s*in|e\s*power|e\s*rev/.test(texto)
      ? 'Importados'
      : 'NP300';
  } else if (modelo === 'KICKS') {
    // Kicks Play e-Power es importado; el Kicks regular es compacto.
    categoria = /play|e\s*power|e\s*rev|(^|\s)power(\s|$)/.test(texto)
      ? 'Importados'
      : 'Compactos';
  } else if (MODELOS_IMPORTADOS.has(modelo)) {
    categoria = 'Importados';
  } else if (MODELOS_COMPACTOS.has(modelo)) {
    categoria = 'Compactos';
  } else if (MODELOS_SUBCOMPACTOS.has(modelo)) {
    categoria = 'Subcompactos';
  }

  return {
    modelo: modelo || 'SIN MODELO',
    categoria,
    clasificacion: CATEGORIA_DETALLE[categoria].letra
  };
}

function extraerCodigoDealer(dealer) {
  const coincidencia = String(dealer ?? '').match(/\d+/);
  return coincidencia ? coincidencia[0] : '';
}

function mapearUbicacion(dealer, sucursal) {
  const textoDealer = normalizarTexto(dealer);
  const codigo = extraerCodigoDealer(dealer);

  if (codigo === '300' && /flotil|flotilla|flo\s*300|(^|\s)flot(\s|$)/.test(textoDealer)) {
    return 'FLOTILLA';
  }

  return UBICACIONES_POR_CODIGO[codigo] || aTexto(sucursal) || 'SIN UBICACIÓN';
}

function detectarSeccion(celdas) {
  if (!Array.isArray(celdas)) return null;

  // En la LISTA DE EXISTENCIAS los marcadores de sección viven en la
  // columna D. No se revisan las demás columnas para no confundir una
  // sucursal llamada «TLALIXTAC» con un cambio de sección.
  const texto = normalizarTexto(celdas[3]);
  if (!texto) return null;

  const seccion = ALIAS_SECCION.get(texto);
  return seccion ? SECCIONES[seccion] : null;
}

module.exports = {
  CATEGORIAS,
  CATEGORIA_DETALLE,
  ORDEN_CATEGORIA,
  UBICACIONES_POR_CODIGO,
  SECCIONES,
  SECCIONES_INCLUIDAS,
  pareceNP300,
  detectarModelo,
  clasificarVehiculo,
  extraerCodigoDealer,
  mapearUbicacion,
  detectarSeccion
};
