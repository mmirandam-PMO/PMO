import fs from 'fs';

const data = JSON.parse(fs.readFileSync('./proyectos.txt', 'utf-8'));
const prog = n => data.programas.find(p => p.nombre === n);

// ---------- Medidas ----------
const W = 1200, M = 30, GAP = 24;
const COL = (W - 2 * M - 2 * GAP) / 3;
const CHAR = 8.2;            // ancho aprox. por caracter (16px)
const LINE = 21;             // alto de línea
const ROW_GAP = 5;
const OWNER_GAP = 28;        // espacio mínimo entre nombre y responsable
const PAD = 16;

const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Colores de texto destacados (como en el diseño original)
const RED = '#C0392B', ORANGE = '#D35400', BLUE = '#2980B9';
const highlight = {
  'Plataforma Banca Empresarial': RED, 'Banca Electrónica y Móvil': RED, 'Hub de pagos': RED,
  'Cuenta Rocket': ORANGE, 'Controles de Integridad': ORANGE, '2FA en trx': ORANGE,
  'Programa Lazos': ORANGE, 'Reingeniería IVR': ORANGE, 'Automatización de procesos ATC': ORANGE,
  'Plataforma Banca Gobierno': BLUE, 'Plataforma Reportes Regulatorios': BLUE,
};
const nameColor = n => {
  const k = Object.keys(highlight).find(k => n.startsWith(k));
  return k ? highlight[k] : '#24292F';
};

function wrap(text, max) {
  const words = text.split(' '); const lines = []; let cur = '';
  for (const w of words) {
    if ((cur + ' ' + w).trim().length > max && cur) { lines.push(cur); cur = w; }
    else cur = (cur + ' ' + w).trim();
  }
  if (cur) lines.push(cur);
  return lines;
}

function row(x, y, name, owner, ownerColor) {
  const tbd = !owner || owner === 'TBD';
  const ownerW = tbd ? 0 : owner.length * 8.6;
  const avail = COL - 2 * PAD - ownerW - (tbd ? 0 : OWNER_GAP);
  const lines = wrap(name, Math.floor(avail / CHAR));
  const h = lines.length * LINE + 22;
  const color = name.startsWith('•') ? '#7F8C8D' : nameColor(name);
  const top = y + (h - lines.length * LINE) / 2 + 15;
  let s = `<rect x="${x}" y="${y}" width="${COL}" height="${h}" rx="5" fill="#F5F7FA"/>`;
  lines.forEach((l, i) => {
    s += `<text x="${x + PAD}" y="${top + i * LINE}" font-size="16" fill="${color}">${esc(l)}</text>`;
  });
  if (!tbd) s += `<text x="${x + COL - PAD}" y="${y + h / 2 + 5}" font-size="14" font-weight="700" text-anchor="end" fill="${ownerColor}">${esc(owner)}</text>`;
  return { svg: s, h };
}

//function header(x, y, title, count, bg, fg = '#FFFFFF') {
  //return `<rect x="${x}" y="${y}" width="${COL}" height="46" rx="7" fill="${bg}"/>` +
   // `<text x="${x + 16}" y="${y + 30}" font-size="19" font-weight="700" fill="${fg}">${esc(title)}</text>` +
    //`<text x="${x + COL - 16}" y="${y + 30}" font-size="19" font-weight="700" text-anchor="end" fill="${fg}">${count}</text>`;
//}

function section(x, y, label, count, rows, ownerColor, rightText) {
  let s = `<text x="${x}" y="${y + 17}" font-size="17" font-weight="700" fill="#24292F">${esc(label)}` +
    (count != null ? `<tspan font-weight="400" fill="#777777" dx="7">• ${count}</tspan>` : '') + `</text>`;
  let off = 0;
  if (rightText) { s += `<text x="${x}" y="${y + 38}" font-size="14" font-weight="700" fill="${ownerColor}">${esc(rightText)}</text>`; off = 21; }
  s += `<line x1="${x}" y1="${y + 28 + off}" x2="${x + COL}" y2="${y + 28 + off}" stroke="#E1E4E8" stroke-width="1"/>`;
  let cy = y + 40 + off;
  for (const [name, owner] of rows) {
    const r = row(x, cy, name, owner, ownerColor);
    s += r.svg; cy += r.h + ROW_GAP;
  }
  return { svg: s, y: cy + 20 };
}

const cat = (p, tipo) => p.categorias.find(c => c.tipo === tipo).proyectos.map(q => [q.nombre, q.responsable]);
const X = i => M + i * (COL + GAP);
const TOP = 24;
let out = '';

// ---------- Columna 1: Banca Minorista ----------
const minor = prog('Banca Minorista');
out += header(X(0), TOP, '🔷  Banca Minorista', minor.total_proyectos, '#4A90E2');
let c1 = section(X(0), TOP + 64, 'Nuevos proyectos', 8, cat(minor, 'Nuevos proyectos'), '#4A90E2');
out += c1.svg;

// ---------- Columna 2: Regulatorios ----------
const reg = prog('Regulatorios');
out += header(X(1), TOP, '⚖️  Regulatorios', reg.total_proyectos, '#A04000');
let c2 = section(X(1), TOP + 64, 'Carryover 2026', 6, cat(reg, 'Carryover 2026'), '#7F8C8D');
out += c2.svg;
let c2b = section(X(1), c2.y, 'Nuevos proyectos', 2, cat(reg, 'Nuevos proyectos'), '#7F8C8D');
out += c2b.svg;

// ---------- Segunda fila alineada: Banca Mayorista | Hub de Innovación ----------
const Y2 = Math.max(c1.y, c2b.y) + 6;

const may = prog('Banca Mayorista');
out += header(X(0), Y2, '🔶  Banca Mayorista', '', '#D35400');
const emp = may.subprogramas[0], gob = may.subprogramas[1], trans = may.subprogramas[2];
const tidy = n => n.replace('Gpos Corporativos', 'Gpos Corp');
const rowsOf = sp => sp.proyectos.map(p => [p.nombre.startsWith('Proyectos ') ? '• ' + p.nombre : tidy(p.nombre), p.responsable]);
let y1 = section(X(0), Y2 + 64, 'Banca de Empresas • Programa', null, rowsOf(emp), '#7F8C8D', emp.responsables.join(' / '));
out += y1.svg;
let y1b = section(X(0), y1.y, 'Banca de Gobierno • Programa', null, rowsOf(gob), '#7F8C8D', gob.responsables.join(' / '));
out += y1b.svg;
let y1c = section(X(0), y1b.y, 'Transversales', 4, rowsOf(trans), '#7F8C8D');
out += y1c.svg;

const hub = prog('Hub de Innovación');
out += header(X(1), Y2, '🔘  Hub de Innovación', hub.total_proyectos, '#BDC3C7', '#555555');
let y2 = section(X(1), Y2 + 64, 'Autorizados', hub.total_proyectos, cat(hub, 'Autorizados'), '#7F8C8D');
out += y2.svg;

// ---------- Columna 3: Carry-over 2026 ----------
const co = prog('Carry-over 2026');
out += header(X(2), TOP, '📂  Carry-over 2026', co.total_proyectos, '#1F3A60');
let cy = TOP + 64, end3 = cy;
for (const c of co.categorias) {
  const r = section(X(2), cy, c.tipo, c.proyectos.length, c.proyectos.map(q => [q.nombre, q.responsable]), '#1F3A60');
  out += r.svg; cy = r.y; end3 = r.y;
}

const H = Math.max(y1c.y, y2.y, end3) + 10;
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="-apple-system, 'Segoe UI', Helvetica, Arial, sans-serif">
<rect width="${W}" height="${H}" fill="#FFFFFF"/>
${out}
</svg>`;
fs.writeFileSync('portafolio.svg', svg);
console.log('portafolio.svg', W, 'x', H);
