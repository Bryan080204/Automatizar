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

const CATEGORIAS_INFO = {
  Nacionales: { clase: 'nacionales', color: '#0f7b53', icono: 'car' },
  Importados: { clase: 'importados', color: '#c3002f', icono: 'ship' },
  Internacionales: { clase: 'internacionales', color: '#2563eb', icono: 'globe' },
  NP300: { clase: 'np300', color: '#b45309', icono: 'truck' },
  'Sin clasificar': { clase: 'sin-clasificar', color: '#687182', icono: 'alert' }
};

const PASOS = [
  'Leyendo el archivo origen...',
  'Detectando encabezados y columnas...',
  'Clasificando en Nacionales, Importados, Internacionales y NP300...',
  'Generando el archivo resultado...'
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
    crearCeldaResumen('Registros', resumen.totalRegistros, {
      destacada: true,
      icono: 'clipboard'
    }),
    ...['Nacionales', 'Importados', 'Internacionales', 'NP300'].map((categoria) => (
      crearCeldaResumen(categoria, conteos.get(categoria) || 0, { categoria })
    ))
  ];

  tarjetas.push(crearCeldaResumen('Para revisión', resumen.paraRevision, {
    destacada: resumen.paraRevision > 0,
    icono: 'alert'
  }));

  elementos.summaryGrid.replaceChildren(...tarjetas);
}

function renderizarCategorias(resultado) {
  const resumen = resultado.resumen;
  const total = resumen.totalRegistros || 0;
  const conteos = new Map((resumen.porCategoria || []).map((item) => [item.label, item.count]));
  const fijas = ['Nacionales', 'Importados', 'Internacionales', 'NP300'];

  elementos.categoryBars.replaceChildren();
  elementos.categoriesMeta.textContent = `4 categorías fijas · ${total} registro(s) en total`;
  elementos.categoriesTotal.textContent = `${total} registro(s)`;

  for (const categoria of fijas) {
    const info = CATEGORIAS_INFO[categoria];
    const count = conteos.get(categoria) || 0;
    const porcentaje = total > 0 ? (count / total) * 100 : 0;
    const fila = document.createElement('div');
    fila.className = 'category-row';
    fila.style.setProperty('--category-color', info.color);

    const etiqueta = document.createElement('span');
    etiqueta.className = `category-badge is-${info.clase}`;
    etiqueta.append(
      crearIcono(info.icono, 'icon category-icon'),
      document.createTextNode(categoria)
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

  const sinClasificar = conteos.get('Sin clasificar') || 0;
  if (sinClasificar > 0) {
    const nota = document.createElement('p');
    nota.className = 'category-note';
    nota.textContent = `${sinClasificar} registro(s) sin clasificar. Se conservaron en la hoja «Sin clasificar» y en la hoja de revisión.`;
    elementos.categoryBars.appendChild(nota);
  }
}

function crearChipDeteccion(etiqueta, valor) {
  const chip = document.createElement('span');
  chip.className = `detect-chip${valor ? '' : ' missing'}`;
  chip.textContent = valor ? `${etiqueta}: ${valor}` : `${etiqueta}: no detectada`;
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

    if (hoja.vacia) {
      const chip = document.createElement('span');
      chip.className = 'detect-chip missing';
      chip.textContent = 'Hoja vacía';
      contenedor.appendChild(chip);
    } else {
      contenedor.append(
        crearChipDeteccion('Tipo', hoja.tipo),
        crearChipDeteccion('Sucursal', hoja.sucursal),
        crearChipDeteccion('Procedencia', hoja.procedencia)
      );
    }

    const registros = document.createElement('small');
    registros.textContent = `${hoja.registros} registro(s)`;
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

  const indiceRevision = resultado.indiceRevision;
  const indiceCategoria = Number.isInteger(resultado.indiceCategoria) ? resultado.indiceCategoria : -1;
  const inicioDerivadas = resultado.inicioDerivadas ?? indiceRevision - 5;
  const indices = resultado.columnas.map((columna, indice) => indice);
  const indicesMostrados = indiceCategoria >= 0
    ? [indiceCategoria, ...indices.filter((indice) => indice !== indiceCategoria)]
    : indices;

  for (const indice of indicesMostrados) {
    elementos.previewHead.appendChild(crearEncabezadoTabla(resultado.columnas[indice]));
  }

  for (const fila of resultado.filas) {
    const tr = document.createElement('tr');

    indicesMostrados.forEach((indice) => {
      const valor = fila[indice];
      const celda = document.createElement('td');
      const texto = valor === null || valor === undefined ? '' : String(valor);
      const esDerivada = indice >= inicioDerivadas && indice <= indiceRevision;

      if (esDerivada) celda.classList.add('is-derived');

      if (texto === '') {
        celda.textContent = '—';
        celda.classList.add('is-empty');
      } else if (indice === indiceCategoria) {
        const info = CATEGORIAS_INFO[texto] || CATEGORIAS_INFO['Sin clasificar'];
        const etiqueta = document.createElement('span');
        etiqueta.className = `category-badge table-badge is-${info.clase}`;
        etiqueta.append(
          crearIcono(info.icono, 'icon table-badge-icon'),
          document.createTextNode(texto)
        );
        celda.classList.add('is-category', `is-${info.clase}`);
        celda.appendChild(etiqueta);
      } else if (indice === indiceRevision) {
        const etiqueta = document.createElement('span');
        etiqueta.className = 'revision-tag';
        etiqueta.append(
          crearIcono('alert', 'icon revision-icon'),
          document.createTextNode('Revisar')
        );
        celda.classList.add('is-revision');
        celda.title = texto;
        celda.appendChild(etiqueta);
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

  elementos.previewMeta.textContent = `${total} registro(s) · 4 categorías · ${resultado.fechaGeneracion}`;
  elementos.previewNote.textContent = mostrados < total
    ? `Mostrando ${mostrados} de ${total} registros. El archivo resultado contiene todos los registros y sus 4 categorías.`
    : `Vista previa completa. El archivo resultado incluye las 4 categorías, el inventario completo y la hoja de revisión.`;
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
  renderizarDeteccion(resultado);
  renderizarTabla(resultado);
  renderizarNotas(resultado);

  const conteos = new Map((resultado.resumen.porCategoria || []).map((item) => [item.label, item.count]));
  const categoriasConRegistros = ['Nacionales', 'Importados', 'Internacionales', 'NP300']
    .filter((categoria) => (conteos.get(categoria) || 0) > 0).length;
  const peso = formatearPeso(resultado.pesoSalida || 0);

  elementos.downloadBtn.href = `/api/descargar/${resultado.token}`;
  elementos.downloadBtn.setAttribute('download', resultado.nombreSalida);
  elementos.downloadFileName.textContent = resultado.nombreSalida;
  elementos.downloadFileMeta.textContent =
    `Excel · ${peso} · ${categoriasConRegistros} de 4 categoría(s) con registros`;
  elementos.downloadDescription.textContent =
    'El archivo incluye las 4 categorías, el inventario completo y la hoja de revisión. El original permanece intacto.';
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
