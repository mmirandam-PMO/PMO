// Tablero ejecutivo "Portafolio de Proyectos 2027" -> PowerPoint editable
// Uso: npm install && npm run build   (genera Portafolio_Proyectos_2027_Tablero.pptx)
const pptxgen = require('pptxgenjs');
const D = require('./data.json');

// ---------- Paleta corporativa ----------
const NAVY = '1B355D', ORANGE = 'E75301', RED = '983920', SKY = '7DACE3', GREY = '8A8F99', GR = '5F6B7F';
const F = 'Calibri', LINE = tint(NAVY, 0.2);
// Mezcla un color con blanco: k=1 color puro, k=0 blanco (tonos claros de la paleta)
function tint(hex, k) {
  return [0, 2, 4].map(i => Math.round(255 - (255 - parseInt(hex.substr(i, 2), 16)) * k)
    .toString(16).padStart(2, '0')).join('').toUpperCase();
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"
const s = pres.addSlide();
s.background = { color: 'F6F8FB' };

// ---------- Helpers ----------
const T = (t, o) => s.addText(t, Object.assign({ fontFace: F, color: NAVY, margin: 0, isTextBox: true, valign: 'middle' }, o));
const rr = (x, y, w, h, fill, line, r) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
  { x, y, w, h, rectRadius: r || 0.06, fill: { color: fill }, line: { color: line || fill, width: 0.75 } });

const HH = 0.72;   // alto de encabezados
const RP = 0.325;  // paso entre filas de proyecto
const RH = 0.295;  // alto de la caja de proyecto
const LB = 0.25;   // alto de la etiqueta de grupo

// Pastilla de responsable
const chip = (x, y, w, h, c, t) => {
  rr(x, y, w, h, tint(c, 0.22), tint(c, 0.22), 0.12);
  T(t, { x, y, w, h, fontSize: 8, bold: true, align: 'center', color: t === '–' ? GR : NAVY });
};
// Caja de proyecto con su responsable
const proj = (x, y, w, c, name, resp) => {
  rr(x, y, w, RH, 'FFFFFF', tint(c, 0.5), 0.05);
  T(name, { x: x + 0.08, y, w: w - 0.95, h: RH, fontSize: 8.5 });
  chip(x + w - 0.82, y + 0.035, 0.74, RH - 0.07, c, resp);
};
const label = (x, y, t) => T(t, { x: x + 0.05, y, w: 3, h: 0.24, fontSize: 10.5, bold: true });
// Encabezado de clasificación con total y dos líneas de resumen
const header = (x, y, w, name, n, c, tc, l1, l2) => {
  rr(x, y, w, HH, c);
  T(name, { x: x + 0.15, y: y + 0.05, w: w - 0.9, h: 0.3, fontSize: 15, bold: true, color: tc });
  T(String(n), { x: x + w - 0.75, y: y + 0.05, w: 0.6, h: 0.3, fontSize: 20, bold: true, color: tc, align: 'right' });
  T(l1, { x: x + 0.15, y: y + 0.37, w: w - 0.3, h: 0.16, fontSize: 8.5, color: tc });
  T(l2, { x: x + 0.15, y: y + 0.53, w: w - 0.3, h: 0.16, fontSize: 8.5, color: tc });
};
// Columna con grupos [etiqueta, [[proyecto, responsable], ...]]
const col = (x, w, y, c, groups) => groups.forEach(([l, items]) => {
  label(x, y, l); y += LB;
  items.forEach(([n, r]) => { proj(x, y, w, c, n, r); y += RP; });
  y += 0.04;
});

// ---------- Título ----------
T(D.titulo, { x: 0.4, y: 0.2, w: 9, h: 0.5, fontSize: 26, bold: true });
T(D.subtitulo, { x: 0.4, y: 0.7, w: 9, h: 0.3, fontSize: 14, bold: true, color: GR });

// ---------- Banca Minorista ----------
const mi = D.minorista;
header(0.4, 1.2, 6.2, mi.nombre, mi.total, SKY, NAVY, mi.linea1, mi.linea2);
label(0.4, 1.95, `${mi.grupo}  ·  ${mi.total}`);
mi.izquierda.forEach((n, i) => proj(0.4, 2.2 + i * RP, 3.0, SKY, n, mi.responsable));
mi.derecha.forEach((n, i) => proj(3.6, 2.2 + i * RP, 3.0, SKY, n, mi.responsable));

// ---------- Banca Mayorista ----------
const my = D.mayorista;
header(0.4, 3.7, 6.2, my.nombre, my.total, ORANGE, 'FFFFFF', my.linea1, my.linea2);
my.programas.forEach((pg, i) => {
  const x = 0.4 + i * 3.2;
  label(x, 4.47, pg.etiqueta);
  rr(x, 4.74, 3.0, 1.2, 'FFFFFF', tint(ORANGE, 0.5), 0.08);
  T(pg.titulo, { x: x + 0.1, y: 4.8, w: 1.9, h: 0.22, fontSize: 9.5, bold: true, color: '2F5EA8' });
  T(pg.detalle, { x: x + 0.1, y: 5.02, w: 2.2, h: 0.2, fontSize: 8, color: GR });
  s.addText(my.bullets.map((b, j) => ({ text: b, options: { bullet: true, breakLine: j < my.bullets.length - 1 } })),
    { x: x + 0.1, y: 5.24, w: 2.6, h: 0.62, fontFace: F, fontSize: 8.5, color: NAVY, margin: 0, valign: 'top', isTextBox: true });
  pg.responsables.forEach((r, j) => chip(x + 2.15, 4.8 + j * 0.28, 0.75, 0.24, ORANGE, r));
});
label(0.4, 6.05, `Transversales  ·  ${my.total}`);
my.transversales.izquierda.forEach(([n, r], i) => proj(0.4, 6.3 + i * RP, 3.0, ORANGE, n, r));
my.transversales.derecha.forEach(([n, r], i) => proj(3.6, 6.3 + i * RP, 3.0, ORANGE, n, r));

// ---------- Regulatorios + Hub de Innovación ----------
const rg = D.regulatorios;
header(6.8, 1.2, 3.0, rg.nombre, rg.total, RED, 'FFFFFF', rg.linea1, rg.linea2);
col(6.8, 3.0, 1.95, RED, rg.grupos);
const hb = D.hub;
header(6.8, 5.25, 3.0, hb.nombre, hb.total, GREY, 'FFFFFF', hb.linea1, hb.linea2);
col(6.8, 3.0, 5.99, GREY, hb.grupos);

// ---------- Carry-over 2026 ----------
const co = D.carryover;
header(10.0, 1.2, 2.93, co.nombre, co.total, NAVY, 'FFFFFF', co.linea1, co.linea2);
col(10.0, 2.93, 1.95, NAVY, co.grupos);

pres.writeFile({ fileName: 'Portafolio_Proyectos_2027_Tablero.pptx' }).then(f => console.log('Generado:', f));
