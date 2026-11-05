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
  toast: document.getElementById('toast')
};

const PASOS = [
  'Leyendo el archivo Excel...',
  'Detectando encabezados y columnas...',
  'Agrupando por tipo, sucursal y procedencia...',
  'Generando el archivo organizado...'
];

const LIMITE_BYTES = 20 * 1024 * 1024;

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
  elementos.fileSize.textContent = `${formatearPeso(archivo.size)} · Listo para organizar`;
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

function crearCeldaResumen(etiqueta, valor, destacada = false) {
  const tarjeta = document.createElement('div');
  tarjeta.className = `summary-card${destacada ? ' highlight' : ''}`;

  const titulo = document.createElement('span');
  titulo.textContent = etiqueta;

  const numero = document.createElement('strong');
  numero.textContent = String(valor);

  tarjeta.append(titulo, numero);
  return tarjeta;
}

function renderizarResumen(resultado) {
  const resumen = resultado.resumen;
  elementos.summaryGrid.replaceChildren(
    crearCeldaResumen('Registros', resumen.totalRegistros, true),
    crearCeldaResumen('Tipos', resumen.tiposUnicos),
    crearCeldaResumen('Sucursales', resumen.sucursalesUnicas),
    crearCeldaResumen('Procedencias', resumen.procedenciasUnicas),
    crearCeldaResumen('Para revisión', resumen.paraRevision, resumen.paraRevision > 0)
  );
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
    nombre.textContent = hoja.nombre;

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

  for (const columna of resultado.columnas) {
    elementos.previewHead.appendChild(crearEncabezadoTabla(columna));
  }

  const indiceRevision = resultado.indiceRevision;
  const inicioDerivadas = indiceRevision - 5;

  for (const fila of resultado.filas) {
    const tr = document.createElement('tr');

    fila.forEach((valor, indice) => {
      const celda = document.createElement('td');
      const texto = valor === null || valor === undefined ? '' : String(valor);
      const esDerivada = indice >= inicioDerivadas && indice <= indiceRevision;

      if (esDerivada) celda.classList.add('is-derived');

      if (texto === '') {
        celda.textContent = '—';
        celda.classList.add('is-empty');
      } else if (indice === indiceRevision) {
        const etiqueta = document.createElement('span');
        etiqueta.className = 'revision-tag';
        etiqueta.textContent = 'Revisar';
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

  elementos.previewMeta.textContent = `${total} registro(s) · ${resultado.hojas.length} hoja(s) · ${resultado.fechaGeneracion}`;
  elementos.previewNote.textContent = mostrados < total
    ? `Mostrando ${mostrados} de ${total} registros. El Excel descargado contiene todos los registros completos.`
    : `Vista previa completa: ${total} registro(s). El Excel descargado contiene la información organizada.`;
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
  renderizarDeteccion(resultado);
  renderizarTabla(resultado);
  renderizarNotas(resultado);

  elementos.downloadBtn.href = `/api/descargar/${resultado.token}`;
  elementos.downloadBtn.setAttribute('download', resultado.nombreSalida);
  elementos.downloadDescription.textContent =
    `Se generó «${resultado.nombreSalida}» con ${resultado.totalRegistros} registro(s). El archivo original permanece intacto.`;
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
      mostrarToast('Archivo procesado correctamente. Tu inventario está listo.');
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
