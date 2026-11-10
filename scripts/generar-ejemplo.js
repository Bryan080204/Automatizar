const path = require('path');
const XLSX = require('xlsx');

const CORTE = Date.UTC(2026, 10, 5);
const RUTA_SALIDA = path.join(__dirname, '..', 'ejemplos', 'inventario-ejemplo.xlsx');

function fechaDesdeDias(dias) {
  const fecha = new Date(CORTE - dias * 86400000);
  const dia = String(fecha.getUTCDate()).padStart(2, '0');
  const mes = String(fecha.getUTCMonth() + 1).padStart(2, '0');
  return `${dia}/${mes}/${fecha.getUTCFullYear()}`;
}

function fila({ sucursal, mod, descripcion, color, serie, dealer, dias, observaciones = '' }) {
  return [
    sucursal,
    '',
    mod,
    descripcion,
    '',
    '',
    color,
    serie,
    '',
    dealer,
    '',
    '',
    fechaDesdeDias(dias),
    dias,
    observaciones,
    ''
  ];
}

function marcador(seccion) {
  return ['', '', '', seccion, '', '', '', '', '', '', '', '', '', '', '', ''];
}

function subtotal(anio, total) {
  return ['SUB-TOTAL', '', '', String(anio), total, '', '', '', '', '', '', '', '', '', '', ''];
}

function totalSeccion(seccion, total) {
  return ['TOTAL', '', '', seccion, total, '', '', '', '', '', '', '', '', '', '', ''];
}

const filas = [
  ['LISTA DE EXISTENCIAS - INVENTARIO NISSAN'],
  ['Corte al', '05/11/2026'],
  [],
  ['UNIDADES POR PAGAR A NISSAN'],
  [],
  ['Sucursal', '', 'Mod', 'Descripcion Local', '', '', 'Colores', 'Serie', '', 'Dealer', '', '', 'Ant.', 'Cod.', 'SISTEMA DE APARTADO', ''],

  marcador('APARTADO'),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2025,
    descripcion: 'VERSA SENSE CVT 1.6 LTS 118 HP',
    color: 'PLATA',
    serie: '601654',
    dealer: ' 300 AVR/Mar',
    dias: 618
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2025,
    descripcion: 'Z TOURING 2 PTAS T/A 3.0 LTS',
    color: 'ROJO SOLIDO BITONO',
    serie: '450005',
    dealer: ' 300 ALLOCAT',
    dias: 525,
    observaciones: 'ENTREGADO-CP-JAVIER'
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2025,
    descripcion: 'URVAN PASAJEROS 4 PUERTAS 11 PASAJEROS T/M 2.5 LTS',
    color: 'BLANCO',
    serie: '121641',
    dealer: '300AVR',
    dias: 433
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2025,
    descripcion: 'MARCH SENSE TM',
    color: 'BLANCO',
    serie: '217591',
    dealer: '287 ALLOCAT',
    dias: 414,
    observaciones: 'REPORTADA EN RCN SOLO CONTADO'
  }),
  subtotal(2026, 4),
  totalSeccion('APARTADO', 4),

  marcador('DEMOSTRACION'),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2026,
    descripcion: 'X-TRAIL 5 PUERTAS ADVANCE CVT 2.5lts 2 ROW',
    color: 'CHAMPAGNE',
    serie: '011125',
    dealer: ' 300 AVR/Mar',
    dias: 341
  }),
  fila({
    sucursal: 'HUAJUAPAN',
    mod: 2025,
    descripcion: 'SENTRA ADVANCE CVT 2.0 LTS',
    color: 'BLANCO PERLADO',
    serie: '425053',
    dealer: '66 AVR/Mar',
    dias: 314,
    observaciones: '5511882214'
  }),
  fila({
    sucursal: 'PUERTO ESCONDIDO',
    mod: 2026,
    descripcion: 'NP300 PICKUP 4X2 T/M 2.5 LTS',
    color: 'GRIS',
    serie: '818758',
    dealer: '298 AVR/Mar',
    dias: 288
  }),
  subtotal(2026, 3),
  totalSeccion('DEMOSTRACION', 3),

  marcador('PISO'),
  fila({
    sucursal: 'MADERO',
    mod: 2025,
    descripcion: 'FRONTIER PLATINUM LE 2.5L 4 CIL TA 4 PTAS 4X2 AA',
    color: 'BLANCO',
    serie: '842973',
    dealer: '287 ALLOCAT',
    dias: 426
  }),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2025,
    descripcion: 'MAGNITE 5 PUERTAS EXCLUSIVE 1.0 LTS T/M TURBO',
    color: 'PLATA',
    serie: '019875',
    dealer: ' 300 AVR/Mar',
    dias: 405,
    observaciones: 'piso accesorios'
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2026,
    descripcion: 'KICKS 5 PUERTAS ADVANCE 2.0 LTS CVT',
    color: 'ROJO BURDEOS',
    serie: '346099',
    dealer: '287 AVR/Mar',
    dias: 310
  }),
  fila({
    sucursal: 'HUATULCO',
    mod: 2026,
    descripcion: 'X-TRAIL 5 PUERTAS PLATINUM CVT 2.5 LTS 2 ROW',
    color: 'GRIS VOLCANICO BITONO',
    serie: '073154',
    dealer: '301 ALLOCAT',
    dias: 310,
    observaciones: '5581665897'
  }),
  fila({
    sucursal: 'PUERTO ESCONDIDO',
    mod: 2026,
    descripcion: 'NP300 PICKUP 4X4 T/M 2.5 LTS',
    color: 'BLANCO',
    serie: '818944',
    dealer: '298 ALLOCAT',
    dias: 279
  }),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2026,
    descripcion: 'VERSA ADVANCE CVT 1.6 LTS 118 HP',
    color: 'BLANCO PERLADO',
    serie: '816513',
    dealer: '300 ALLOCAT',
    dias: 257
  }),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2026,
    descripcion: 'KICKS PLAY e-POWER ELECTRICO 5 PUERTAS PLATINUM MOTOR',
    color: 'CHAMPAGNE',
    serie: '011256',
    dealer: '300 AVR/Mar',
    dias: 255
  }),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2026,
    descripcion: 'FRONTIER V6 PRO-4X T/A 4X4 MOTOR 3.8 LTS 6 CIL 4P',
    color: 'VERDE MITICO',
    serie: '633070',
    dealer: '300 ALLOCAT',
    dias: 241,
    observaciones: 'ACCESORIOS $15,387.81'
  }),
  fila({
    sucursal: 'HUAJUAPAN',
    mod: 2026,
    descripcion: 'SENTRA SR PLATINUM CVT 2.0 LTS',
    color: 'NARANJA MAGMA BITONO',
    serie: '237434',
    dealer: '66 AVR/Mar',
    dias: 219,
    observaciones: 'DEMO 66 AVR/Mar'
  }),
  fila({
    sucursal: 'HUATULCO',
    mod: 2027,
    descripcion: 'KAIT 5 PUERTAS PLATINUM 1.6 LTS CVT',
    color: 'BLANCO PERLADO',
    serie: '100700',
    dealer: '301 AVR/Mar',
    dias: 191
  }),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2026,
    descripcion: 'PATHFINDER 5 PUERTAS PLATINUM 3.5 LTS 4X4',
    color: 'NEGRO',
    serie: '710245',
    dealer: '300 ALLOCAT',
    dias: 120
  }),
  fila({
    sucursal: 'HUAJUAPAN',
    mod: 2026,
    descripcion: 'V-DRIVE 4 PUERTAS T/M A/C AUDIO 1.6 LTS',
    color: 'BLANCO',
    serie: '414939',
    dealer: '66 ALLOCAT',
    dias: 95
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2026,
    descripcion: 'SENTRA ADVANCE T/M 2.0 LTS',
    color: 'GRIS',
    serie: '811248',
    dealer: '',
    dias: 60
  }),
  fila({
    sucursal: 'PUEBLO NUEVO',
    mod: 2026,
    descripcion: 'NP300 PICKUP 4X2 T/M 2.5 LTS',
    color: 'BLANCO',
    serie: '806426',
    dealer: ' DEMO 300 FLOTIL',
    dias: 45,
    observaciones: 'FLOT MUNICIPIO DE SAN JUAN TEITIPAC'
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2027,
    descripcion: 'Z TOURING 2 PTAS T/A 3.0 LTS',
    color: 'AMARILLO',
    serie: '450018',
    dealer: '287 AVR/Mar',
    dias: 12
  }),
  subtotal(2025, 2),
  subtotal(2026, 10),
  subtotal(2027, 3),
  totalSeccion('PISO', 15),

  marcador('SEMI_NUEVOS GARANTIZADOS'),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2024,
    descripcion: 'SENTRA ADVANCE CVT 2.0 LTS',
    color: 'NARANJA SOLAR',
    serie: '268582',
    dealer: '300 ALLOCAT',
    dias: 563
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2024,
    descripcion: 'X-TRAIL 5 PUERTAS EXCLUSIVE CVT 2.5 LTS 2 ROW',
    color: 'AZUL IMPERIAL',
    serie: '012395',
    dealer: '287 AVR/Mar',
    dias: 349
  }),
  subtotal(2024, 2),
  totalSeccion('SEMI_NUEVOS GARANTIZADOS', 2),

  marcador('SEMINUEVOS'),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2023,
    descripcion: 'VERSA EXCLUSIVE CVT 1.6',
    color: 'NARANJA SOLAR',
    serie: '70835',
    dealer: '300 ALLOCAT',
    dias: 770
  }),
  fila({
    sucursal: 'HUAJUAPAN',
    mod: 2022,
    descripcion: 'MARCH ADVANCE T/M',
    color: 'AZUL',
    serie: '209373',
    dealer: '66 AVR/Mar',
    dias: 66
  }),
  subtotal(2023, 2),
  totalSeccion('SEMINUEVOS', 2),

  marcador('TLALIXTAC'),
  fila({
    sucursal: 'TLALIXTAC',
    mod: 2023,
    descripcion: 'NP300 PICKUP 4X2 T/M 2.5 LTS',
    color: 'GRIS',
    serie: '777001',
    dealer: '301 ALLOCAT',
    dias: 500
  }),
  fila({
    sucursal: 'TLALIXTAC',
    mod: 2022,
    descripcion: 'VERSA SENSE T/M 1.6 LTS 118 HP',
    color: 'BLANCO',
    serie: '777002',
    dealer: '300 AVR/Mar',
    dias: 420
  }),
  subtotal(2023, 2),
  totalSeccion('TLALIXTAC', 2),

  marcador('USADOS'),
  fila({
    sucursal: 'UNIVERSIDAD',
    mod: 2024,
    descripcion: 'RENAULT LOGAN 1.6 INTENS MT',
    color: 'BLANCO GLACIAR',
    serie: '533114',
    dealer: '300 ALLOCAT',
    dias: 253
  }),
  fila({
    sucursal: 'MADERO',
    mod: 2021,
    descripcion: 'SENTRA ADVANCE T/M 2.0 LTS',
    color: 'GRIS OXFORD',
    serie: '392820',
    dealer: '287 ALLOCAT',
    dias: 800
  }),
  subtotal(2024, 2),
  totalSeccion('USADOS', 2),

  ['TOTAL:', '', '', '', 30, '', '', '', '', '', '', '', '', '', '', ''],
  [],
  ['El archivo original no se modifica', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
];

const hoja = XLSX.utils.aoa_to_sheet(filas);
hoja['!cols'] = [
  { wch: 24 }, { wch: 2 }, { wch: 7 }, { wch: 54 }, { wch: 2 }, { wch: 2 },
  { wch: 24 }, { wch: 14 }, { wch: 2 }, { wch: 20 }, { wch: 2 }, { wch: 2 },
  { wch: 12 }, { wch: 8 }, { wch: 38 }, { wch: 2 }
];
hoja['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } },
  { s: { r: 1, c: 0 }, e: { r: 1, c: 0 } },
  { s: { r: 3, c: 0 }, e: { r: 3, c: 8 } }
];

const libro = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(libro, hoja, 'LISTA DE EXISTENCIAS');
XLSX.writeFile(libro, RUTA_SALIDA);

console.log(`Ejemplo generado: ${RUTA_SALIDA}`);
console.log(`Unidades incluidas: 22 · excluidas: 8 · total: 30`);
