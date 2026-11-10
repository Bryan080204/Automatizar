const elementos = {
  workspace: document.getElementById('workspace'),
  uploadSection: document.getElementById('uploadSection'),
  dropzone: document.getElementById('dropzone'),
  fileInput: document.getElementById('fileInput'),
  fileCard: document.getElementById('fileCard'),
  fileName: document.getElementById('fileName'),
  fileSize: document.getElementById('fileSize'),
  removeFile: document.getElementById('removeFile'),
  processBtn: document.getElementById('processBtn'),
  processing: document.getElementById('processing'),
  processingStep: document.getElementById('processingStep'),
  progressBar: document.getElementById('progressBar'),
  results: document.getElementById('results'),
  resetBtn: document.getElementById('resetBtn'),
  summaryGrid: document.getElementById('summaryGrid'),
  categoriesMeta: document.getElementById('categoriesMeta'),
  categoriesTotal: document.getElementById('categoriesTotal'),
  categoryBars: document.getElementById('categoryBars'),
  locationsMeta: document.getElementById('locationsMeta'),
  locationsTotal: document.getElementById('locationsTotal'),
  locationsHead: document.getElementById('locationsHead'),
  locationsBody: document.getElementById('locationsBody'),
  sheetsProcessed: document.getElementById('sheetsProcessed'),
  detectionList: document.getElementById('detectionList'),
  previewHead: document.getElementById('previewHead'),
  previewBody: document.getElementById('previewBody'),
  previewMeta: document.getElementById('previewMeta'),
  previewNote: document.getElementById('previewNote'),
  warningsPanel: document.getElementById('warningsPanel'),
  warningsList: document.getElementById('warningsList'),
  downloadBtn: document.getElementById('downloadBtn'),
  downloadDescription: document.getElementById('downloadDescription'),
  downloadFileName: document.getElementById('downloadFileName'),
  downloadFileMeta: document.getElementById('downloadFileMeta'),
  toast: document.getElementById('toast')
};

const CATEGORIAS = ['Importados', 'Compactos', 'NP300', 'Subcompactos'];

const CATEGORIAS_INFO = {
  Importados: { clase: 'importados', color: '#c3002f', icono: 'ship', letra: 'I' },
  Compactos: { clase: 'compactos', color: '#2563eb', icono: 'car', letra: 'C' },
  NP300: { clase: 'np300', color: '#b45309', icono: 'truck', letra: 'D' },
  Subcompactos: { clase: 'subcompactos', color: '#0f7b53', icono: 'small-car', letra: 'S' },
  'Sin clasificar': { clase: 'sin-clasificar', color: '#687182', icono: 'alert', letra: '?' }
};

const PASOS = [
  'Leyendo el archivo origen...',
  'Ubicando encabezados y secciones...',
  'Limpiando seminuevos y clasificando unidades...',
  'Generando resumen, ubicaciones y hojas...'
];

const LIMITE_BYTES = 20 * 1024 * 1024;

function crearIcono(nombre, clase = 'icon') {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', clase);
  svg.setAttribute('aria-hidden', 'true');
  const uso = document.createElementNS('http://www.w3.org/2000/svg', 'use');
  uso.setAttribute('href', `#i-${nombre}`);
  svg.appendChild(uso);
  return svg;
}

let archivoSeleccionado = null;
let resultadoActual = null;
let temporizadorProgreso = null;
let temporizadorToast = null;

function formatearPeso(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function mostrarToast(mensaje, tipo = 'success') {
  clearTimeout(temporizadorToast);
  elementos.toast.textContent = mensaje;
  elementos.toast.classList.toggle('error', tipo === 'error');
  elementos.toast.classList.remove('is-hidden');

  temporizadorToast = setTimeout(() => {
    elementos.toast.classList.add('is-hidden');
  }, tipo === 'error' ? 6500 : 4500);
}

function validarArchivo(archivo) {
  if (!archivo) return false;

  const extension = archivo.name.split('.').pop().toLowerCase();

  if (extension !== 'xlsx' && extension !== 'xls') {
    mostrarToast('Solo se permiten archivos de Excel con extensión .xlsx o .xls.', 'error');
    return false;
  }

  if (archivo.size > LIMITE_BYTES) {
    mostrarToast('El archivo supera el límite de 20 MB.', 'error');
    return false;
  }

  return true;
}

function seleccionarArchivo(archivo) {
  if (!validarArchivo(archivo)) {
    elementos.fileInput.value = '';
    return;
  }

  archivoSeleccionado = archivo;
  elementos.fileName.textContent = archivo.name;
  elementos.fileSize.textContent = `${formatearPeso(archivo.size)} · Listo para generar el resultado`;
  elementos.fileCard.classList.remove('is-hidden');
  elementos.processBtn.disabled = false;
  elementos.processing.classList.add('is-hidden');
}

function limpiarArchivo() {
  archivoSeleccionado = null;
  elementos.fileInput.value = '';
  elementos.fileCard.classList.add('is-hidden');
  elementos.fileName.textContent = '';
  elementos.fileSize.textContent = '';
  elementos.processBtn.disabled = true;
  elementos.processing.classList.add('is-hidden');
  elementos.progressBar.style.width = '0%';
}

function iniciarProgreso() {
  clearInterval(temporizadorProgreso);
  elementos.processing.classList.remove('is-hidden');
  elementos.processBtn.disabled = true;
  elementos.progressBar.style.width = '8%';
  elementos.processingStep.textContent = PASOS[0];

  let paso = 0;

  temporizadorProgreso = setInterval(() => {
    paso = Math.min(paso + 1, PASOS.length - 1);
    elementos.processingStep.textContent = PASOS[paso];
    elementos.progressBar.style.width = `${Math.min(86, 18 + paso * 22)}%`;
  }, 700);
}

function finalizarProgreso() {
  clearInterval(temporizadorProgreso);
  elementos.processingStep.textContent = 'Archivo listo.';
  elementos.progressBar.style.width = '100%';
}

function detenerProgreso() {
  clearInterval(temporizadorProgreso);
  elementos.processing.classList.add('is-hidden');
  elementos.progressBar.style.width = '0%';
  elementos.processBtn.disabled = !archivoSeleccionado;
}

function crearCeldaResumen(etiqueta, valor, opciones = {}) {
  const { destacada = false, categoria = null, icono = null } = opciones;
  const info = categoria ? (CATEGORIAS_INFO[categoria] || CATEGORIAS_INFO['Sin clasificar']) : null;
  const tarjeta = document.createElement('div');
  tarjeta.className = `summary-card${destacada ? ' highlight' : ''}${categoria ? ' category-card' : ''}`;

  const titulo = document.createElement('span');
  titulo.textContent = etiqueta;

  if (info) {
    tarjeta.classList.add(`is-${info.clase}`);
    tarjeta.style.setProperty('--category-color', info.color);
  }

  const numero = document.createElement('strong');
  numero.textContent = String(valor);

  const iconoTarjeta = document.createElement('span');
  iconoTarjeta.className = 'summary-icon';
  iconoTarjeta.appendChild(crearIcono(icono || (info ? info.icono : 'clipboard')));

  tarjeta.append(titulo, numero, iconoTarjeta);
  return tarjeta;
}

function renderizarResumen(resultado) {
  const resumen = resultado.resumen;
  const conteos = new Map((resumen.porCategoria || []).map((item) => [item.label, item.count]));
  const tarjetas = [
    crearCeldaResumen('Unidades incluidas', resumen.totalRegistros, {
      destacada: true,
      icono: 'clipboard'
    }),
    ...CATEGORIAS.map((categoria) => (
      crearCeldaResumen(categoria, conteos.get(categoria) || 0, { categoria })
    )),
    crearCeldaResumen('Excluidas', resumen.totalExcluidos || 0, {
      icono: 'alert'
    })
  ];

  if ((resumen.sinClasificar || 0) > 0) {
    tarjetas.push(crearCeldaResumen('Sin clasificar', resumen.sinClasificar, {
      categoria: 'Sin clasificar'
    }));
  }

  elementos.summaryGrid.replaceChildren(...tarjetas);
}

function renderizarCategorias(resultado) {
  const resumen = resultado.resumen;
  const total = resumen.totalRegistros || 0;
  const excluidos = resumen.totalExcluidos || 0;
  const conteos = new Map((resumen.porCategoria || []).map((item) => [item.label, item.count]));

  elementos.categoryBars.replaceChildren();
  elementos.categoriesMeta.textContent = `4 categorías · ${total} incluidas · ${excluidos} excluidas`;
  elementos.categoriesTotal.textContent = `${total} unidades`;

  for (const categoria of CATEGORIAS) {
    const info = CATEGORIAS_INFO[categoria];
    const count = conteos.get(categoria) || 0;
    const porcentaje = total > 0 ? (count / total) * 100 : 0;
    const fila = document.createElement('div');
    fila.className = 'category-row';
    fila.style.setProperty('--category-color', info.color);

    const etiqueta = document.createElement('span');
    etiqueta.className = `category-badge is-${info.clase}`;

    const clave = document.createElement('span');
    clave.className = 'category-key';
    clave.textContent = info.letra;

    etiqueta.append(
      crearIcono(info.icono, 'icon category-icon'),
      document.createTextNode(categoria),
      clave
    );

    const medidor = document.createElement('div');
    medidor.className = 'category-meter';
    const barra = document.createElement('span');
    barra.style.width = `${porcentaje}%`;
    medidor.appendChild(barra);

    const cantidad = document.createElement('strong');
    cantidad.textContent = String(count);

    const proporcion = document.createElement('small');
    proporcion.textContent = `${porcentaje.toFixed(0)}%`;

    fila.append(etiqueta, medidor, cantidad, proporcion);
    elementos.categoryBars.appendChild(fila);
  }

  const seccionesExcluidas = (resumen.secciones || [])
    .filter((seccion) => !seccion.incluida && seccion.count > 0);

  if (excluidos > 0) {
    const nota = document.createElement('p');
    nota.className = 'category-note';
    nota.textContent = `Se excluyeron ${excluidos} unidad(es) de ${seccionesExcluidas
      .map((seccion) => `${seccion.label} (${seccion.count})`)
      .join(', ')}.`;
    elementos.categoryBars.appendChild(nota);
  }

  const sinClasificar = conteos.get('Sin clasificar') || 0;
  if (sinClasificar > 0) {
    const nota = document.createElement('p');
    nota.className = 'category-note neutral';
    nota.textContent = `${sinClasificar} unidad(es) no coinciden con el catálogo de modelos. Se conservaron en la hoja «SIN CLASIFICAR».`;
    elementos.categoryBars.appendChild(nota);
  }
}

function renderizarUbicaciones(resultado) {
  const resumen = resultado.resumen;
  const ubicaciones = resumen.porUbicacion || [];
  const total = resumen.totalRegistros || 0;

  elementos.locationsHead.replaceChildren();
  elementos.locationsBody.replaceChildren();

  if (!ubicaciones.length) {
    elementos.locationsMeta.textContent = 'Sin ubicaciones detectadas';
    elementos.locationsTotal.textContent = '0 ubicaciones';
    return;
  }

  const haySinClasificar = ubicaciones.some(
    (item) => (item.categorias && item.categorias['Sin clasificar']) > 0
  );
  const categorias = [...CATEGORIAS, ...(haySinClasificar ? ['Sin clasificar'] : [])];
  const columnas = ['Ubicación', 'Unidades', '% del total', ...categorias];

  for (const columna of columnas) {
    const th = document.createElement('th');
    th.textContent = columna;
    const info = CATEGORIAS_INFO[columna];

    if (info) {
      th.classList.add(`is-${info.clase}`);
      th.style.setProperty('--category-color', info.color);
    }

    elementos.locationsHead.appendChild(th);
  }

  const totalesCategoria = new Map(categorias.map((categoria) => [categoria, 0]));

  for (const ubicacion of ubicaciones) {
    const fila = document.createElement('tr');
    const porcentaje = total > 0 ? (ubicacion.count / total) * 100 : 0;

    const celdaNombre = document.createElement('td');
    const nombre = document.createElement('span');
    nombre.className = 'location-name';
    const marcador = document.createElement('span');
    marcador.className = 'location-marker';
    nombre.append(marcador, document.createTextNode(ubicacion.label));
    celdaNombre.appendChild(nombre);
    fila.appendChild(celdaNombre);

    const celdaUnidades = document.createElement('td');
    celdaUnidades.className = 'location-units';
    celdaUnidades.textContent = String(ubicacion.count);
    fila.appendChild(celdaUnidades);

    const celdaPorcentaje = document.createElement('td');
    celdaPorcentaje.className = 'location-share-value';
    celdaPorcentaje.textContent = `${porcentaje.toFixed(1)}%`;
    fila.appendChild(celdaPorcentaje);

    for (const categoria of categorias) {
      const count = (ubicacion.categorias && ubicacion.categorias[categoria]) || 0;
      totalesCategoria.set(categoria, (totalesCategoria.get(categoria) || 0) + count);

      const celda = document.createElement('td');
      celda.className = 'location-category';
      const info = CATEGORIAS_INFO[categoria] || CATEGORIAS_INFO['Sin clasificar'];
      const punto = document.createElement('span');
      punto.className = `location-dot is-${info.clase}`;
      const valor = document.createElement('span');
      valor.textContent = String(count);
      celda.append(punto, valor);
      fila.appendChild(celda);
    }

    elementos.locationsBody.appendChild(fila);
  }

  const filaTotal = document.createElement('tr');
  filaTotal.className = 'total-row';

  const totalNombre = document.createElement('td');
  totalNombre.textContent = 'TOTAL GLOBAL';
  filaTotal.appendChild(totalNombre);

  const totalUnidades = document.createElement('td');
  totalUnidades.textContent = String(total);
  filaTotal.appendChild(totalUnidades);

  const totalPorcentaje = document.createElement('td');
  totalPorcentaje.className = 'location-share-value';
  totalPorcentaje.textContent = '100.0%';
  filaTotal.appendChild(totalPorcentaje);

  for (const categoria of categorias) {
    const celda = document.createElement('td');
    celda.className = 'location-category';
    celda.textContent = String(totalesCategoria.get(categoria) || 0);
    filaTotal.appendChild(celda);
  }

  elementos.locationsBody.appendChild(filaTotal);

  elementos.locationsMeta.textContent =
    `${ubicaciones.length} ubicaciones de origen · ${total} unidades · mezcla por categoría`;
  elementos.locationsTotal.textContent = `${ubicaciones.length} ubicaciones`;
}

function crearChipDeteccion(texto, clase = '') {
  const chip = document.createElement('span');
  chip.className = `detect-chip${clase ? ` ${clase}` : ''}`;
  chip.textContent = texto;
  return chip;
}

function renderizarDeteccion(resultado) {
  elementos.detectionList.replaceChildren();
  elementos.sheetsProcessed.textContent = `${resultado.hojas.length} hoja(s) analizada(s)`;

  for (const hoja of resultado.hojas) {
    const contenedor = document.createElement('div');
    contenedor.className = 'detection-sheet';

    const nombre = document.createElement('strong');
    nombre.append(
      crearIcono('file-sheet', 'icon detection-icon'),
      document.createTextNode(hoja.nombre)
    );

    contenedor.appendChild(nombre);

    if (!hoja.valida) {
      contenedor.appendChild(crearChipDeteccion('Hoja omitida', 'missing'));
      if (hoja.motivo) contenedor.appendChild(crearChipDeteccion(hoja.motivo, 'missing'));
    } else {
      if (hoja.encabezadoFila) {
        contenedor.appendChild(crearChipDeteccion(`Encabezado: fila ${hoja.encabezadoFila}`));
      }

      for (const seccion of hoja.secciones || []) {
        contenedor.appendChild(
          crearChipDeteccion(
            `${seccion.label}: ${seccion.count}`,
            seccion.incluida ? 'included' : 'excluded'
          )
        );
      }
    }

    const registros = document.createElement('small');
    registros.textContent = hoja.valida
      ? `${hoja.registros} incluidas · ${hoja.excluidos} excluidas`
      : 'Sin datos para el reporte';
    contenedor.appendChild(registros);

    elementos.detectionList.appendChild(contenedor);
  }
}

function crearEncabezadoTabla(texto) {
  const celda = document.createElement('th');
  celda.textContent = texto;
  celda.title = texto;
  return celda;
}

function renderizarTabla(resultado) {
  elementos.previewHead.replaceChildren();
  elementos.previewBody.replaceChildren();

  const indiceCategoria = Number.isInteger(resultado.indiceCategoria) ? resultado.indiceCategoria : -1;
  const indiceClasificacion = Number.isInteger(resultado.indiceClasificacion) ? resultado.indiceClasificacion : -1;
  const indiceDias = resultado.columnas.findIndex((columna) => columna === 'Días');
  const indices = resultado.columnas.map((columna, indice) => indice);
  const preferidos = [indiceCategoria, indiceClasificacion, indiceDias].filter((indice) => indice >= 0);
  const indicesMostrados = [...new Set([...preferidos, ...indices])];

  for (const indice of indicesMostrados) {
    elementos.previewHead.appendChild(crearEncabezadoTabla(resultado.columnas[indice]));
  }

  for (const fila of resultado.filas) {
    const tr = document.createElement('tr');
    const textoCategoria = indiceCategoria >= 0 ? String(fila[indiceCategoria] ?? '') : '';
    const infoCategoria = CATEGORIAS_INFO[textoCategoria] || CATEGORIAS_INFO['Sin clasificar'];

    indicesMostrados.forEach((indice) => {
      const valor = fila[indice];
      const celda = document.createElement('td');
      const texto = valor === null || valor === undefined ? '' : String(valor);

      if (texto === '') {
        celda.textContent = '—';
        celda.classList.add('is-empty');
      } else if (indice === indiceCategoria) {
        const etiqueta = document.createElement('span');
        etiqueta.className = `category-badge table-badge is-${infoCategoria.clase}`;
        etiqueta.append(
          crearIcono(infoCategoria.icono, 'icon table-badge-icon'),
          document.createTextNode(texto)
        );
        celda.classList.add('is-category', `is-${infoCategoria.clase}`);
        celda.appendChild(etiqueta);
      } else if (indice === indiceClasificacion) {
        const clave = document.createElement('span');
        clave.className = 'table-key';
        clave.textContent = texto;
        celda.classList.add('is-key', `is-${infoCategoria.clase}`);
        celda.title = `${texto} · ${textoCategoria}`;
        celda.appendChild(clave);
      } else if (indice === indiceDias) {
        const dias = Number(texto);
        celda.textContent = texto;
        celda.classList.add('is-days');
        if (Number.isFinite(dias) && dias >= 365) {
          celda.classList.add('is-old');
        } else if (Number.isFinite(dias) && dias >= 180) {
          celda.classList.add('is-aged');
        }
      } else {
        celda.textContent = texto;
        if (texto.length > 32) celda.title = texto;
      }

      tr.appendChild(celda);
    });

    elementos.previewBody.appendChild(tr);
  }

  const total = resultado.totalRegistros;
  const mostrados = resultado.filasMostradas;
  const ubicaciones = (resultado.resumen.porUbicacion || []).length;

  elementos.previewMeta.textContent =
    `${total} unidades · 4 categorías · ${ubicaciones} ubicaciones · ordenado por Días`;
  elementos.previewNote.textContent = mostrados < total
    ? `Mostrando ${mostrados} de ${total} unidades. El archivo resultado contiene la lista completa y las hojas por categoría.`
    : 'Vista previa completa. El archivo resultado incluye resumen ejecutivo, lista completa, hojas por categoría y hojas por ubicación de origen.';
}

function renderizarNotas(resultado) {
  const notas = [...resultado.advertencias, ...resultado.notasInterpretacion];
  elementos.warningsList.replaceChildren();

  if (!notas.length) {
    elementos.warningsPanel.classList.add('is-hidden');
    return;
  }

  for (const nota of notas) {
    const item = document.createElement('li');
    item.textContent = nota;
    elementos.warningsList.appendChild(item);
  }

  elementos.warningsPanel.classList.remove('is-hidden');
}

function renderizarResultado(resultado) {
  resultadoActual = resultado;

  renderizarResumen(resultado);
  renderizarCategorias(resultado);
  renderizarUbicaciones(resultado);
  renderizarDeteccion(resultado);
  renderizarTabla(resultado);
  renderizarNotas(resultado);

  const conteos = new Map((resultado.resumen.porCategoria || []).map((item) => [item.label, item.count]));
  const categoriasConRegistros = CATEGORIAS
    .filter((categoria) => (conteos.get(categoria) || 0) > 0).length;
  const ubicaciones = (resultado.resumen.porUbicacion || []).length;
  const peso = formatearPeso(resultado.pesoSalida || 0);

  elementos.downloadBtn.href = `/api/descargar/${resultado.token}`;
  elementos.downloadBtn.setAttribute('download', resultado.nombreSalida);
  elementos.downloadFileName.textContent = resultado.nombreSalida;
  elementos.downloadFileMeta.textContent =
    `Excel · ${peso} · ${categoriasConRegistros} de 4 categorías · ${ubicaciones} ubicaciones`;
  elementos.downloadDescription.textContent =
    'Incluye resumen ejecutivo, distribución por ubicación, lista completa ordenada por antigüedad, hojas por categoría y hojas por ubicación de origen. El original permanece intacto.';
}

async function procesarArchivo() {
  if (!archivoSeleccionado) {
    mostrarToast('Selecciona un archivo Excel antes de continuar.', 'error');
    return;
  }

  const formulario = new FormData();
  formulario.append('archivo', archivoSeleccionado);

  iniciarProgreso();

  try {
    const respuesta = await fetch('/api/procesar', {
      method: 'POST',
      body: formulario
    });

    const datos = await respuesta.json().catch(() => ({}));

    if (!respuesta.ok) {
      throw new Error(datos.error || 'No se pudo procesar el archivo.');
    }

    finalizarProgreso();

    setTimeout(() => {
      elementos.workspace.classList.add('is-hidden');
      elementos.results.classList.remove('is-hidden');
      renderizarResultado(datos);
      elementos.results.scrollIntoView({ behavior: 'smooth', block: 'start' });
      detenerProgreso();
      mostrarToast('Archivo procesado. Tu archivo resultado está listo para descargar.');
    }, 450);
  } catch (error) {
    detenerProgreso();
    mostrarToast(error.message || 'Ocurrió un error al procesar el archivo.', 'error');
  }
}

function reiniciar() {
  resultadoActual = null;
  limpiarArchivo();
  elementos.results.classList.add('is-hidden');
  elementos.workspace.classList.remove('is-hidden');
  elementos.downloadBtn.href = '#';
  elementos.workspace.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

elementos.dropzone.addEventListener('click', (evento) => {
  if (evento.target === elementos.fileInput) return;
  elementos.fileInput.click();
});

elementos.dropzone.addEventListener('keydown', (evento) => {
  if (evento.key === 'Enter' || evento.key === ' ') {
    evento.preventDefault();
    elementos.fileInput.click();
  }
});

elementos.fileInput.addEventListener('change', () => {
  const archivo = elementos.fileInput.files && elementos.fileInput.files[0];
  if (archivo) seleccionarArchivo(archivo);
});

elementos.removeFile.addEventListener('click', () => {
  limpiarArchivo();
});

elementos.processBtn.addEventListener('click', procesarArchivo);
elementos.resetBtn.addEventListener('click', reiniciar);

elementos.downloadBtn.addEventListener('click', () => {
  if (elementos.downloadBtn.getAttribute('href') === '#') {
    mostrarToast('Primero procesa un archivo para generar la descarga.', 'error');
    return;
  }
  mostrarToast('Descarga iniciada.');
});

for (const evento of ['dragenter', 'dragover']) {
  elementos.dropzone.addEventListener(evento, (e) => {
    e.preventDefault();
    e.stopPropagation();
    elementos.dropzone.classList.add('is-dragging');
  });
}

for (const evento of ['dragleave', 'drop']) {
  elementos.dropzone.addEventListener(evento, (e) => {
    e.preventDefault();
    e.stopPropagation();
    elementos.dropzone.classList.remove('is-dragging');
  });
}

elementos.dropzone.addEventListener('drop', (evento) => {
  const archivo = evento.dataTransfer.files && evento.dataTransfer.files[0];
  if (archivo) seleccionarArchivo(archivo);
});

window.addEventListener('dragover', (evento) => evento.preventDefault());
window.addEventListener('drop', (evento) => evento.preventDefault());
