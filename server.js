const path = require('path');
const crypto = require('crypto');
const express = require('express');
const multer = require('multer');
const { procesarInventario, ErrorInventario } = require('./src/inventoryProcessor');

const app = express();
const PORT = process.env.PORT || 3000;
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const EXPIRACION = 30 * 60 * 1000;
const resultados = new Map();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_FILE_SIZE, files: 1 },
  fileFilter(req, file, callback) {
    const extension = path.extname(file.originalname || '').toLowerCase();
    if (extension === '.xlsx' || extension === '.xls') {
      callback(null, true);
      return;
    }
    callback(new Error('Solo se permiten archivos .xlsx o .xls.'));
  }
});

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '1mb' }));

app.post('/api/procesar', upload.single('archivo'), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'Selecciona un archivo Excel para procesar.' });
      return;
    }

    const resultado = await procesarInventario(req.file.buffer, req.file.originalname);
    const token = crypto.randomUUID();

    resultados.set(token, {
      buffer: resultado.buffer,
      nombreSalida: resultado.nombreSalida,
      expira: Date.now() + EXPIRACION
    });

    res.json({ ...resultado.preview, token });
  } catch (error) {
    const estado = error instanceof ErrorInventario ? error.estado : 500;
    const mensaje = error instanceof ErrorInventario
      ? error.message
      : 'Ocurrió un error al procesar el archivo. Intenta de nuevo.';

    if (!(error instanceof ErrorInventario)) console.error(error);
    res.status(estado).json({ error: mensaje });
  }
});

app.get('/api/descargar/:token', (req, res) => {
  const resultado = resultados.get(req.params.token);

  if (!resultado) {
    res.status(404).json({ error: 'El enlace de descarga no existe o expiró. Procesa el archivo nuevamente.' });
    return;
  }

  if (resultado.expira <= Date.now()) {
    resultados.delete(req.params.token);
    res.status(410).json({ error: 'El enlace de descarga expiró. Procesa el archivo nuevamente.' });
    return;
  }

  const nombreAscii = resultado.nombreSalida
    .replace(/[^\x20-\x7E]/g, '_')
    .replace(/["\\]/g, '_');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader(
    'Content-Disposition',
    `attachment; filename="${nombreAscii}"; filename*=UTF-8''${encodeURIComponent(resultado.nombreSalida)}`
  );
  res.send(resultado.buffer);
});

setInterval(() => {
  const ahora = Date.now();
  for (const [token, resultado] of resultados) {
    if (resultado.expira <= ahora) resultados.delete(token);
  }
}, 5 * 60 * 1000).unref();

app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      res.status(413).json({ error: 'El archivo supera el límite de 20 MB.' });
      return;
    }
    res.status(400).json({ error: `Error al recibir el archivo: ${error.message}` });
    return;
  }

  if (error) {
    if (error.message === 'Solo se permiten archivos .xlsx o .xls.') {
      res.status(400).json({ error: error.message });
      return;
    }
    console.error(error);
    res.status(500).json({ error: 'Ocurrió un error inesperado.' });
    return;
  }

  next();
});

app.listen(PORT, () => {
  console.log(`Organizador de inventarios Nissan disponible en http://localhost:${PORT}`);
});
