// Portafolio de Proyectos 2027 -> PowerPoint editable (7 diapositivas)
//   1) Tablero   2) Presupuesto por clasificación   3-7) Timeline + presupuesto por clasificación
// Uso: npm install && npm run build   (genera Portafolio_Proyectos_2027.pptx)
const pptxgen = require('pptxgenjs');
const D = require('./data.json');

// ---------- Paleta corporativa ----------
const NAVY = '1B355D', ORANGE = 'E75301', RED = '983920', SKY = '7DACE3', GREY = '8A8F99', GR = '5F6B7F';
const COLOR = { NAVY, ORANGE, RED, SKY, GREY };                       // nombre usado en data.json -> hex
const DIA = { NAVY, ORANGE: 'C24400', RED, SKY: '2F67B8', GREY: GR };  // color fuerte para rombos/números
const F = 'Calibri';
// Mezcla un color con blanco: k=1 color puro, k=0 blanco (tonos claros de la paleta)
function tint(hex, k) {
  return [0, 2, 4].map(i => Math.round(255 - (255 - parseInt(hex.substr(i, 2), 16)) * k)
    .toString(16).padStart(2, '0')).join('').toUpperCase();
}
const LINE = tint(NAVY, 0.2);

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"

// Helpers ligados a una diapositiva
function kit(s) {
  return {
    T: (t, o) => s.addText(t, Object.assign({ fontFace: F, color: NAVY, margin: 0, isTextBox: true, valign: 'middle' }, o)),
    rr: (x, y, w, h, fill, line, r) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
      { x, y, w, h, rectRadius: r || 0.06, fill: { color: fill }, line: { color: line || fill, width: 0.75 } })
  };
}
const newSlide = () => { const s = pres.addSlide(); s.background = { color: 'F6F8FB' }; return s; };

// Preparación de clasificaciones (total de proyectos y líneas)
// Fila: [proyecto, línea, responsable, plan_ini, ejec_ini, ejec_fin, peso?]  (meses 0=Ene ... 11=Dic)
// peso = cuántos proyectos representa la fila (por defecto 1; 0 para programas TBD)
const CLS = D.clasificaciones.map(k => ({
  nombre: k.nombre, c: COLOR[k.color], key: k.color, filas: k.filas,
  n: k.filas.reduce((s, r) => s + (r[6] === undefined ? 1 : r[6]), 0),
  lineas: [...new Set(k.filas.map(r => r[1]))]
}));

/* =====================================================================
   DIAPOSITIVA 1 · TABLERO
   ===================================================================== */
function tablero() {
  const s = newSlide(); const { T, rr } = kit(s);
  const HH = 0.72, RP = 0.325, RH = 0.295, LB = 0.25;
  const chip = (x, y, w, h, c, t) => {
    rr(x, y, w, h, tint(c, 0.22), tint(c, 0.22), 0.12);
    T(t, { x, y, w, h, fontSize: 8, bold: true, align: 'center', color: t === '–' ? GR : NAVY });
  };
  const proj = (x, y, w, c, name, resp) => {
    rr(x, y, w, RH, 'FFFFFF', tint(c, 0.5), 0.05);
    T(name, { x: x + 0.08, y, w: w - 0.95, h: RH, fontSize: 8.5 });
    chip(x + w - 0.82, y + 0.035, 0.74, RH - 0.07, c, resp);
  };
  const label = (x, y, t) => T(t, { x: x + 0.05, y, w: 3, h: 0.24, fontSize: 10.5, bold: true });
  const header = (x, y, w, name, n, c, tc, l1, l2) => {
    rr(x, y, w, HH, c);
    T(name, { x: x + 0.15, y: y + 0.05, w: w - 0.9, h: 0.3, fontSize: 15, bold: true, color: tc });
    T(String(n), { x: x + w - 0.75, y: y + 0.05, w: 0.6, h: 0.3, fontSize: 20, bold: true, color: tc, align: 'right' });
    T(l1, { x: x + 0.15, y: y + 0.37, w: w - 0.3, h: 0.16, fontSize: 8.5, color: tc });
    T(l2, { x: x + 0.15, y: y + 0.53, w: w - 0.3, h: 0.16, fontSize: 8.5, color: tc });
  };
  const col = (x, w, y, c, groups) => groups.forEach(([l, items]) => {
    label(x, y, l); y += LB;
    items.forEach(([n, r]) => { proj(x, y, w, c, n, r); y += RP; });
    y += 0.04;
  });

  T(D.titulo, { x: 0.4, y: 0.2, w: 9, h: 0.5, fontSize: 26, bold: true });
  T(D.subtitulo, { x: 0.4, y: 0.7, w: 9, h: 0.3, fontSize: 14, bold: true, color: GR });

  const mi = D.minorista;
  header(0.4, 1.2, 6.2, mi.nombre, mi.total, SKY, NAVY, mi.linea1, mi.linea2);
  label(0.4, 1.95, `${mi.grupo}  ·  ${mi.total}`);
  mi.izquierda.forEach((n, i) => proj(0.4, 2.2 + i * RP, 3.0, SKY, n, mi.responsable));
  mi.derecha.forEach((n, i) => proj(3.6, 2.2 + i * RP, 3.0, SKY, n, mi.responsable));

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

  const rg = D.regulatorios;
  header(6.8, 1.2, 3.0, rg.nombre, rg.total, RED, 'FFFFFF', rg.linea1, rg.linea2);
  col(6.8, 3.0, 1.95, RED, rg.grupos);
  const hb = D.hub;
  header(6.8, 5.25, 3.0, hb.nombre, hb.total, GREY, 'FFFFFF', hb.linea1, hb.linea2);
  col(6.8, 3.0, 5.99, GREY, hb.grupos);

  const co = D.carryover;
  header(10.0, 1.2, 2.93, co.nombre, co.total, NAVY, 'FFFFFF', co.linea1, co.linea2);
  col(10.0, 2.93, 1.95, NAVY, co.grupos);
}

/* =====================================================================
   DIAPOSITIVA 2 · PRESUPUESTO POR CLASIFICACIÓN (sin total general)
   ===================================================================== */
function presupuesto() {
  const s = newSlide(); const { T, rr } = kit(s);
  T('Presupuesto 2027 por clasificación', { x: 0.4, y: 0.2, w: 11, h: 0.5, fontSize: 26, bold: true });
  T('Un espacio por clasificación · sin total general · captura el monto en cada tarjeta', { x: 0.4, y: 0.7, w: 11, h: 0.3, fontSize: 14, bold: true, color: GR });
  const W = 2.37, GAP = 0.17;
  CLS.forEach((k, i) => {
    const x = 0.4 + i * (W + GAP), y = 1.3;
    rr(x, y, W, 5.4, 'FFFFFF', LINE, 0.08);
    rr(x, y, W, 0.9, tint(k.c, 0.3), tint(k.c, 0.3), 0.08);
    T(k.nombre, { x: x + 0.15, y: y + 0.1, w: W - 0.3, h: 0.4, fontSize: 15, bold: true });
    T(`${k.n} proyectos`, { x: x + 0.15, y: y + 0.5, w: W - 0.3, h: 0.3, fontSize: 11 });
    T('PRESUPUESTO 2027', { x: x + 0.15, y: y + 1.1, w: W - 0.3, h: 0.25, fontSize: 9.5, bold: true, color: GR });
    rr(x + 0.15, y + 1.4, W - 0.3, 0.75, tint(k.c, 0.1), tint(k.c, 0.45), 0.08);
    T('$ –', { x: x + 0.15, y: y + 1.4, w: W - 0.3, h: 0.75, fontSize: 28, bold: true, align: 'center', color: DIA[k.key] });
    T('LÍNEAS', { x: x + 0.15, y: y + 2.4, w: W - 0.3, h: 0.25, fontSize: 9.5, bold: true, color: GR });
    k.lineas.forEach((l, j) => {
      const cnt = k.filas.filter(r => r[1] === l).reduce((a, r) => a + (r[6] === undefined ? 1 : r[6]), 0);
      const yy = y + 2.7 + j * 0.42;
      T(l, { x: x + 0.15, y: yy, w: W - 0.85, h: 0.36, fontSize: 10.5 });
      rr(x + W - 0.65, yy + 0.05, 0.5, 0.26, tint(k.c, 0.22), tint(k.c, 0.22), 0.12);
      T(cnt ? String(cnt) : 'TBD', { x: x + W - 0.65, y: yy + 0.05, w: 0.5, h: 0.26, fontSize: 9.5, bold: true, align: 'center' });
    });
  });
  T('Fechas y montos por capturar · el presupuesto por proyecto se registra en cada diapositiva de timeline.', { x: 0.4, y: 6.95, w: 12.5, h: 0.25, fontSize: 10, italic: true, color: GR });
  s.addNotes('Escribe el monto en el cuadro "$ –" de cada clasificación. No hay total general por diseño.');
}

/* =====================================================================
   DIAPOSITIVAS 3-7 · TIMELINE + PRESUPUESTO (una fila por proyecto)
   ===================================================================== */
function timeline(k) {
  const s = newSlide(); const { T, rr } = kit(s);
  const MX = 7.05, MW = 0.49, X0 = 0.4;                      // inicio y ancho de cada mes
  const kp = k.key === 'NAVY' ? 0.2 : 0.3, ke = k.key === 'NAVY' ? 0.42 : 0.6;  // tonos planeación / ejecución
  T(k.nombre, { x: 0.4, y: 0.2, w: 8, h: 0.5, fontSize: 26, bold: true });
  T(`${k.n} proyectos · ${k.lineas.length} ${k.lineas.length > 1 ? 'líneas' : 'línea'} · Timeline y presupuesto 2027`, { x: 0.4, y: 0.7, w: 8.5, h: 0.3, fontSize: 14, bold: true, color: GR });
  // Presupuesto de la clasificación
  rr(9.9, 0.25, 3.03, 0.75, tint(k.c, 0.25), tint(k.c, 0.5), 0.08);
  T('PRESUPUESTO DE LA CLASIFICACIÓN', { x: 10.05, y: 0.3, w: 2.8, h: 0.22, fontSize: 8.5, bold: true });
  T('$ –', { x: 10.05, y: 0.52, w: 2.8, h: 0.42, fontSize: 22, bold: true, color: DIA[k.key] });
  // Encabezados de columnas y meses
  [['Proyecto', X0 + 0.1, 2.9], ['Línea', 3.4, 1.3], ['Responsable', 4.8, 1.15], ['Presupuesto', 5.95, 1.0]]
    .forEach(([t, x, w], i) => T(t, { x, y: 1.35, w, h: 0.34, fontSize: 10, bold: true, color: GR, align: i === 3 ? 'center' : 'left' }));
  D.meses.forEach((m, i) => {
    rr(MX + i * MW + 0.02, 1.35, MW - 0.04, 0.34, tint(NAVY, 0.12), tint(NAVY, 0.12), 0.05);
    T(m, { x: MX + i * MW, y: 1.35, w: MW, h: 0.34, fontSize: 10, bold: true, align: 'center' });
  });
  // Filas
  const n = k.filas.length, H = Math.min(0.56, 4.85 / n), big = H >= 0.45, fs = big ? 10 : 8.5;
  k.filas.forEach((r, i) => {
    const y = 1.85 + i * H, h = H - 0.06;
    rr(X0, y, 12.53, h, 'FFFFFF', LINE, 0.06);
    T(r[0], { x: X0 + 0.1, y, w: 2.85, h, fontSize: fs, bold: true });
    T(r[1], { x: 3.4, y, w: 1.35, h, fontSize: fs - 0.5, color: GR });
    T(r[2], { x: 4.8, y, w: 1.15, h, fontSize: fs - 0.5 });
    rr(5.98, y + (h - 0.26) / 2, 0.94, 0.26, 'FFFFFF', tint(k.c, 0.5), 0.12);        // presupuesto del proyecto
    T('$ –', { x: 5.98, y: y + (h - 0.26) / 2, w: 0.94, h: 0.26, fontSize: 10, align: 'center', color: GR });
    rr(MX, y + 0.04, 12 * MW, h - 0.08, tint(NAVY, 0.04), tint(NAVY, 0.04), 0.05);   // pista
    const bh = Math.min(0.2, h - 0.14), by = y + (h - bh) / 2, [ps, es, ee] = [r[3], r[4], r[5]];
    if (es > ps) rr(MX + ps * MW + 0.03, by, (es - ps) * MW - 0.06, bh, tint(k.c, kp), tint(k.c, kp), bh / 2);   // planeación
    if (ee >= es) {
      rr(MX + es * MW + 0.03, by, (ee - es + 1) * MW - 0.06, bh, tint(k.c, ke), tint(k.c, ke), bh / 2);          // ejecución
      if (ee < 11) s.addShape(pres.shapes.DIAMOND, { x: MX + (ee + 1) * MW - 0.11, y: y + (h - 0.22) / 2, w: 0.22, h: 0.22, fill: { color: DIA[k.key] }, line: { color: 'FFFFFF', width: 1 } }); // hito
    }
  });
  // Leyenda
  const LY = 1.85 + n * H + 0.02;
  rr(0.4, LY + 0.05, 0.34, 0.16, tint(k.c, kp), tint(k.c, kp), 0.08); T('Planeación', { x: 0.82, y: LY, w: 1.1, h: 0.26, fontSize: 10, color: GR });
  rr(2.0, LY + 0.05, 0.34, 0.16, tint(k.c, ke), tint(k.c, ke), 0.08); T('Ejecución', { x: 2.42, y: LY, w: 1.1, h: 0.26, fontSize: 10, color: GR });
  s.addShape(pres.shapes.DIAMOND, { x: 3.65, y: LY + 0.03, w: 0.2, h: 0.2, fill: { color: DIA[k.key] }, line: { color: 'FFFFFF', width: 1 } });
  T('Hito / salida a producción', { x: 3.95, y: LY, w: 2.3, h: 0.26, fontSize: 10, color: GR });
  T('Fechas ilustrativas · presupuesto por capturar', { x: 8.4, y: LY, w: 4.53, h: 0.26, fontSize: 10, italic: true, color: GR, align: 'right' });
  s.addNotes('Para actualizar: arrastra o estira las barras (planeación/ejecución) y el rombo de hito sobre los meses. Escribe el presupuesto en cada "$ –" y el de la clasificación arriba a la derecha.');
}

// ---------- Armado ----------
tablero();
presupuesto();
CLS.forEach(timeline);
pres.writeFile({ fileName: 'Portafolio_Proyectos_2027.pptx' }).then(f => console.log('Generado:', f));
