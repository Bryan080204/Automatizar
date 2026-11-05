const MAX_COLUMNAS = 120;
const MAX_FILAS_ENCABEZADO = 30;

const PISTAS_ENCABEZADO = new Set([
  'marca', 'modelo', 'version', 'vin', 'serie', 'color', 'placa', 'precio',
  'tipo', 'sucursal', 'sucursales', 'agencia', 'agencias', 'procedencia',
  'ciudad', 'origen', 'descripcion', 'clave', 'stock', 'inventario',
  'vehiculo', 'vehiculos', 'unidad', 'unidades', 'anio', 'ano', 'numero',
  'no', 'codigo', 'id', 'estatus', 'fecha', 'motor', 'transmision',
  'combustible', 'puertas', 'carroceria', 'pais', 'localidad', 'municipio',
  'region', 'zona', 'plaza', 'concesionario', 'distribuidor', 'tienda',
  'dealer', 'kilometraje', 'km', 'condicion', 'categoria', 'clasificacion',
  'nacional', 'importado', 'internacional', 'extranjero', 'estado', 'entidad',
  'nombre', 'cantidad', 'total', 'concepto', 'importe', 'folio', 'cliente',
  'observaciones', 'notas', 'comentarios'
]);

const PISTAS_NACIONALIDAD = new Set([
  'nacional', 'nacionales', 'importado', 'importados', 'internacional',
  'internacionales', 'extranjero', 'extranjeros', 'local', 'locales',
  'nacionalizado', 'nacionalizada', 'importacion', 'importaciones'
]);

function normalizarTexto(valor) {
  return String(valor ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function contieneFrase(textoNormalizado, frase) {
  return ` ${textoNormalizado} `.includes(` ${frase} `);
}

function aTexto(valor) {
  if (valor === null || valor === undefined) return '';

  if (valor instanceof Date) {
    if (Number.isNaN(valor.getTime())) return '';
    const dia = String(valor.getDate()).padStart(2, '0');
    const mes = String(valor.getMonth() + 1).padStart(2, '0');
    const anio = valor.getFullYear();
    const hora = valor.getHours();
    const minutos = valor.getMinutes();
    if (hora || minutos) {
      return `${dia}/${mes}/${anio} ${String(hora).padStart(2, '0')}:${String(minutos).padStart(2, '0')}`;
    }
    return `${dia}/${mes}/${anio}`;
  }

  if (typeof valor === 'number') {
    return Number.isFinite(valor) ? String(valor) : '';
  }

  if (typeof valor === 'boolean') {
    return valor ? 'Sí' : 'No';
  }

  return String(valor).trim();
}

function esVacio(valor) {
  return valor === null || valor === undefined || (typeof valor === 'string' && valor.trim() === '');
}

function pareceTextoDeEncabezado(valor) {
  const texto = aTexto(valor);
  if (!texto) return false;
  if (!/[a-zA-Z\u00C0-\u024F]/.test(texto)) return false;
  return texto.length <= 80;
}

function contarPistasEncabezado(valor) {
  const normalizado = normalizarTexto(valor);
  if (!normalizado) return 0;
  return normalizado
    .split(' ')
    .filter((token) => PISTAS_ENCABEZADO.has(token))
    .length;
}

function detectarFilaEncabezado(filas) {
  const limite = Math.min(filas.length, MAX_FILAS_ENCABEZADO);
  let mejor = null;

  for (let i = 0; i < limite; i += 1) {
    const fila = Array.isArray(filas[i]) ? filas[i] : [];
    const celdas = fila.map(aTexto).filter(Boolean);
    if (celdas.length < 2) continue;

    let pistas = 0;
    let textos = 0;
    let numericos = 0;

    for (const celda of celdas) {
      pistas += contarPistasEncabezado(celda);
      if (pareceTextoDeEncabezado(celda)) {
        textos += 1;
      } else {
        numericos += 1;
      }
    }

    if (pistas === 0) continue;

    const puntaje = pistas * 8 + textos * 3 - numericos * 2 + Math.min(celdas.length, 20) * 0.25;
    if (!mejor || puntaje > mejor.puntaje) {
      mejor = { indice: i, puntaje };
    }
  }

  if (mejor) return mejor.indice;

  let mejorSinPistas = null;

  for (let i = 0; i < Math.min(filas.length, 12); i += 1) {
    const fila = Array.isArray(filas[i]) ? filas[i] : [];
    const celdas = fila.map(aTexto).filter(Boolean);
    if (celdas.length < 2) continue;

    const textos = celdas.filter(pareceTextoDeEncabezado).length;
    const proporcionTexto = textos / celdas.length;
    if (proporcionTexto < 0.6) continue;

    const puntaje = celdas.length + proporcionTexto * 3;

    if (!mejorSinPistas || puntaje > mejorSinPistas.puntaje) {
      mejorSinPistas = { indice: i, puntaje };
    }
  }

  return mejorSinPistas ? mejorSinPistas.indice : -1;
}

function calcularTotalColumnas(filas, filaEncabezado) {
  let ultima = -1;

  const revisar = (fila) => {
    if (!Array.isArray(fila)) return;
    const limite = Math.min(fila.length, MAX_COLUMNAS);
    for (let i = 0; i < limite; i += 1) {
      if (!esVacio(fila[i])) ultima = Math.max(ultima, i);
    }
  };

  revisar(filaEncabezado || []);
  for (const fila of filas) revisar(fila);

  return Math.max(ultima + 1, 1);
}

function construirEncabezados(filaEncabezado, totalColumnas) {
  const encabezados = [];
  const generadas = [];
  const repetidas = [];
  const vistos = new Map();

  for (let i = 0; i < totalColumnas; i += 1) {
    const original = aTexto(Array.isArray(filaEncabezado) ? filaEncabezado[i] : '');
    const generada = !original;
    const labelBase = original || `Columna ${i + 1}`;
    const base = normalizarTexto(labelBase) || `columna ${i + 1}`;
    const ocurrencia = (vistos.get(base) || 0) + 1;
    vistos.set(base, ocurrencia);

    if (generada) generadas.push(i);
    if (ocurrencia > 1 && !generada) repetidas.push(labelBase);

    const label = ocurrencia > 1 && !generada ? `${labelBase} (${ocurrencia})` : labelBase;

    encabezados.push({
      index: i,
      label,
      labelBase,
      base,
      ocurrencia,
      generada
    });
  }

  return { encabezados, generadas, repetidas };
}

module.exports = {
  PISTAS_NACIONALIDAD,
  normalizarTexto,
  contieneFrase,
  aTexto,
  esVacio,
  contarPistasEncabezado,
  detectarFilaEncabezado,
  calcularTotalColumnas,
  construirEncabezados
};
