const path = require('path');
const XLSX = require('xlsx');

const filasGeneral = [
  ['INVENTARIO VEHÍCULOS NISSAN - CORTE OCTUBRE 2026'],
  ['Reporte interno de unidades disponibles'],
  [],
  ['No.', 'Sucursal', 'Procedencia', 'Tipo', 'Marca', 'Modelo', 'Año', 'Color', 'VIN', 'Precio', 'Fecha ingreso'],
  [1, 'Madero', 'Huajuapan', 'Nacional', 'Nissan', 'Versa', 2024, 'Blanco', '3N1CN7AP4RL800101', 285000, new Date(2026, 9, 1)],
  [2, 'Universidad', 'Oaxaca capital', 'Importado', 'Nissan', 'Kicks', 2023, 'Gris', '3N8CP5HE2PL800102', 320000, new Date(2026, 9, 2)],
  [3, 'Madero', 'Puerto Escondido', 'Nacional', 'Nissan', 'Sentra', 2024, 'Rojo', '3N1AB8CV5RY800103', 345000, new Date(2026, 9, 3)],
  [4, '', 'Huajuapan', 'Nacional', 'Nissan', 'March', 2022, 'Azul', '3N1CK3CP7NL800104', 210000, new Date(2026, 9, 4)],
  [5, 'Universidad', '', '', 'Nissan', 'X-Trail', 2021, 'Negro', 'JN8AT3BB5MW800105', 390000, new Date(2026, 9, 5)],
  [6, 'Centro', 'Oaxaca capital', 'Internacional', 'Nissan', 'Pathfinder', 2024, 'Plata', '5N1DR3CC1RC800106', 620000, new Date(2026, 9, 6)],
  [7, 'Madero', 'Huajuapan', 'Nacional', 'Nissan', 'Versa', 2023, 'Blanco', '', 260000, new Date(2026, 9, 7)],
  ['Total de unidades: 7']
];

const hojaGeneral = XLSX.utils.aoa_to_sheet(filasGeneral);
hojaGeneral['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },
  { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
  { s: { r: 11, c: 0 }, e: { r: 11, c: 10 } }
];
hojaGeneral['!cols'] = [
  { wch: 7 }, { wch: 15 }, { wch: 19 }, { wch: 14 }, { wch: 11 },
  { wch: 13 }, { wch: 8 }, { wch: 11 }, { wch: 21 }, { wch: 12 }, { wch: 15 }
];

const filasImportados = [
  ['UNIDADES IMPORTADAS'],
  [],
  ['VIN', 'Modelo', 'Marca', 'Año', 'Color', 'Origen', 'Ciudad', 'Sucursal', 'Precio'],
  ['3N8CP5HE2PL800201', 'Kicks', 'Nissan', 2024, 'Gris', 'Importado', 'Oaxaca capital', 'Universidad', 330000],
  ['JN8AT3BB5MW800202', 'X-Trail', 'Nissan', 2023, 'Negro', 'Internacional', 'Huajuapan', 'Madero', 410000],
  ['3N1AB8CV5RY800203', 'Sentra', 'Nissan', 2022, 'Blanco', 'Importado', 'Puerto Escondido', 'Centro', 295000]
];

const hojaImportados = XLSX.utils.aoa_to_sheet(filasImportados);
hojaImportados['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }
];
hojaImportados['!cols'] = [
  { wch: 21 }, { wch: 13 }, { wch: 11 }, { wch: 8 }, { wch: 11 },
  { wch: 15 }, { wch: 19 }, { wch: 15 }, { wch: 12 }
];

const libro = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(libro, hojaGeneral, 'Inventario general');
XLSX.utils.book_append_sheet(libro, hojaImportados, 'Importados');

const ruta = path.join(__dirname, '..', 'ejemplos', 'inventario-ejemplo.xlsx');
XLSX.writeFile(libro, ruta);

console.log(`Ejemplo creado en: ${ruta}`);
