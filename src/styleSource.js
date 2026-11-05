const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');

function esXlsx(buffer) {
  return buffer.length > 4
    && buffer[0] === 0x50
    && buffer[1] === 0x4b
    && buffer[2] === 0x03
    && buffer[3] === 0x04;
}

function esArchivoOle(buffer) {
  return buffer.length > 4
    && buffer[0] === 0xd0
    && buffer[1] === 0xcf
    && buffer[2] === 0x11
    && buffer[3] === 0xe0;
}

function convertirConExcel(rutaEntrada, rutaSalida, rutaScript) {
  const script = [
    "$ErrorActionPreference = 'Stop'",
    '$excel = New-Object -ComObject Excel.Application',
    '$excel.Visible = $false',
    '$excel.DisplayAlerts = $false',
    'try {',
    `  $libro = $excel.Workbooks.Open('${rutaEntrada.replace(/'/g, "''")}')`,
    `  $libro.SaveAs('${rutaSalida.replace(/'/g, "''")}', 51)`,
    '  $libro.Close($false)',
    '} finally {',
    '  $excel.Quit()',
    '}'
  ].join('\r\n');

  fs.writeFileSync(rutaScript, script, 'utf8');

  return new Promise((resolve, reject) => {
    execFile(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', rutaScript],
      { timeout: 90000, windowsHide: true },
      (error) => {
        if (error) {
          reject(error);
          return;
        }
        resolve();
      }
    );
  });
}

async function obtenerXlsxConFormato(buffer) {
  if (!buffer || !buffer.length) return null;
  if (esXlsx(buffer)) return buffer;
  if (!esArchivoOle(buffer)) return null;

  const carpeta = fs.mkdtempSync(path.join(os.tmpdir(), 'inventario-nissan-'));
  const entrada = path.join(carpeta, 'origen.xls');
  const salida = path.join(carpeta, 'origen.xlsx');
  const rutaScript = path.join(carpeta, 'convertir.ps1');

  try {
    fs.writeFileSync(entrada, buffer);
    await convertirConExcel(entrada, salida, rutaScript);

    if (!fs.existsSync(salida)) return null;
    return fs.readFileSync(salida);
  } catch (error) {
    return null;
  } finally {
    fs.rm(carpeta, { recursive: true, force: true }, () => {});
  }
}

module.exports = {
  obtenerXlsxConFormato
};
