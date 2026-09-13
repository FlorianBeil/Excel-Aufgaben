/* Excel.Flo – Symbolbibliothek für den Excel-/Power-Query-Nachbau
 *
 * Eigene, vereinfachte SVG-Symbole im Stil der Office-Menübänder (keine
 * Microsoft-Grafiken). Große Symbole: 32×32, kleine: 16×16.
 * Aufruf: ExcelFloIcons.get("removeColumns", 32) → SVG-Markup.
 */

(function () {
  "use strict";

  const C = {
    g: "#217346", // Excel-Grün
    gd: "#185C37",
    gl: "#CDE8D7",
    b: "#2B7CD3",
    bl: "#CFE3F7",
    r: "#D13438",
    o: "#D86C1E",
    y: "#E8A70E",
    p: "#7A5CC7",
    k: "#5E5E5E",
    kl: "#A6A6A6",
    line: "#C8C8C8",
    w: "#FFFFFF",
  };

  function svg(size, body) {
    return (
      '<svg class="xl-ico xl-ico--' +
      size +
      '" width="' +
      size +
      '" height="' +
      size +
      '" viewBox="0 0 ' +
      size +
      " " +
      size +
      '" aria-hidden="true" focusable="false">' +
      body +
      "</svg>"
    );
  }

  /* ---------------- Grundformen (32er-Raster) ---------------- */

  function table32(o) {
    o = o || {};
    const x = o.x == null ? 3.5 : o.x;
    const y = o.y == null ? 6.5 : o.y;
    const w = o.w || 25;
    const h = o.h || 20;
    const cols = o.cols || 3;
    const rows = o.rows || 4;
    const rh = h / rows;
    const cw = w / cols;
    let s = '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="#fff" stroke="' + C.k + '"/>';
    if (o.selCol != null) {
      s += '<rect x="' + (x + cw * o.selCol) + '" y="' + y + '" width="' + cw + '" height="' + h + '" fill="' + (o.selFill || C.gl) + '"/>';
    }
    if (o.selRow != null) {
      s += '<rect x="' + x + '" y="' + (y + rh * o.selRow) + '" width="' + w + '" height="' + rh + '" fill="' + (o.selFill || C.gl) + '"/>';
    }
    s += '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + rh + '" fill="' + (o.hdr || C.g) + '"/>';
    for (let i = 1; i < cols; i++) {
      s += '<line x1="' + (x + cw * i) + '" y1="' + (y + rh) + '" x2="' + (x + cw * i) + '" y2="' + (y + h) + '" stroke="' + C.line + '"/>';
    }
    for (let j = 2; j < rows; j++) {
      s += '<line x1="' + x + '" y1="' + (y + rh * j) + '" x2="' + (x + w) + '" y2="' + (y + rh * j) + '" stroke="' + C.line + '"/>';
    }
    s += '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="none" stroke="' + C.k + '"/>';
    return s;
  }

  function table16(o) {
    o = o || {};
    const x = o.x == null ? 1.5 : o.x;
    const y = o.y == null ? 2.5 : o.y;
    const w = o.w || 13;
    const h = o.h || 11;
    let s = '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="#fff" stroke="' + C.k + '"/>';
    if (o.selCol != null) s += '<rect x="' + (x + (w / 3) * o.selCol) + '" y="' + y + '" width="' + w / 3 + '" height="' + h + '" fill="' + C.gl + '"/>';
    s += '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="3" fill="' + (o.hdr || C.g) + '"/>';
    s += '<line x1="' + (x + w / 3) + '" y1="' + (y + 3) + '" x2="' + (x + w / 3) + '" y2="' + (y + h) + '" stroke="' + C.line + '"/>';
    s += '<line x1="' + (x + (2 * w) / 3) + '" y1="' + (y + 3) + '" x2="' + (x + (2 * w) / 3) + '" y2="' + (y + h) + '" stroke="' + C.line + '"/>';
    s += '<line x1="' + x + '" y1="' + (y + 7) + '" x2="' + (x + w) + '" y2="' + (y + 7) + '" stroke="' + C.line + '"/>';
    s += '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="none" stroke="' + C.k + '"/>';
    return s;
  }

  function doc32(fill) {
    return (
      '<path d="M7.5 3.5h12l6 6v19h-18z" fill="' +
      (fill || "#fff") +
      '" stroke="' +
      C.k +
      '"/><path d="M19.5 3.5v6h6" fill="none" stroke="' +
      C.k +
      '"/>'
    );
  }

  function doc16(fill) {
    return '<path d="M3.5 1.5h6l3 3v10h-9z" fill="' + (fill || "#fff") + '" stroke="' + C.k + '"/><path d="M9.5 1.5v3h3" fill="none" stroke="' + C.k + '"/>';
  }

  function xMark(cx, cy, r, sw) {
    return (
      '<path d="M' + (cx - r) + " " + (cy - r) + "L" + (cx + r) + " " + (cy + r) + "M" + (cx + r) + " " + (cy - r) + "L" + (cx - r) + " " + (cy + r) +
      '" stroke="' + C.r + '" stroke-width="' + (sw || 2.4) + '" stroke-linecap="round"/>'
    );
  }

  function check(cx, cy, s, color) {
    return (
      '<path d="M' + (cx - s) + " " + cy + "l" + s * 0.7 + " " + s * 0.7 + "l" + s * 1.3 + " " + -s * 1.4 +
      '" fill="none" stroke="' + (color || C.g) + '" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>'
    );
  }

  function plus(cx, cy, r, color, sw) {
    return (
      '<path d="M' + (cx - r) + " " + cy + "H" + (cx + r) + "M" + cx + " " + (cy - r) + "V" + (cy + r) +
      '" stroke="' + (color || C.g) + '" stroke-width="' + (sw || 2.4) + '" stroke-linecap="round"/>'
    );
  }

  function refresh(cx, cy, r, color, sw) {
    const c = color || C.g;
    return (
      '<path d="M' + (cx + r) + " " + cy + "A" + r + " " + r + " 0 1 1 " + (cx + r * 0.3) + " " + (cy - r * 0.95) +
      '" fill="none" stroke="' + c + '" stroke-width="' + (sw || 2.2) + '"/>' +
      '<path d="M' + (cx + r * 0.1) + " " + (cy - r * 1.55) + "l" + r * 0.75 + " " + r * 0.6 + "l" + -r * 0.85 + " " + r * 0.45 + 'z" fill="' + c + '"/>'
    );
  }

  function funnel32(color) {
    return '<path d="M4.5 5.5h23l-9 11v9l-5 3v-12z" fill="' + (color || "#fff") + '" stroke="' + C.k + '" stroke-linejoin="round"/>';
  }

  function funnel16(color) {
    return '<path d="M1.5 2.5h12l-4.5 5.5v5l-3 2v-7z" fill="' + (color || "#fff") + '" stroke="' + C.k + '" stroke-linejoin="round"/>';
  }

  function gear(cx, cy, r, color) {
    const c = color || C.k;
    let teeth = "";
    for (let i = 0; i < 8; i++) {
      const a = (Math.PI / 4) * i;
      teeth +=
        '<line x1="' + (cx + Math.cos(a) * r * 0.7) + '" y1="' + (cy + Math.sin(a) * r * 0.7) + '" x2="' + (cx + Math.cos(a) * r) +
        '" y2="' + (cy + Math.sin(a) * r) + '" stroke="' + c + '" stroke-width="' + r * 0.38 + '" stroke-linecap="round"/>';
    }
    return teeth + '<circle cx="' + cx + '" cy="' + cy + '" r="' + r * 0.62 + '" fill="' + c + '"/><circle cx="' + cx + '" cy="' + cy + '" r="' + r * 0.27 + '" fill="#fff"/>';
  }

  function text(x, y, str, size, color, weight) {
    return (
      '<text x="' + x + '" y="' + y + '" font-family="Segoe UI, Arial, sans-serif" font-size="' + size + '" font-weight="' + (weight || 700) +
      '" fill="' + (color || C.k) + '">' + str + "</text>"
    );
  }

  function sortArrow16(desc, color) {
    return desc
      ? '<path d="M12 2.5v10M9.5 10l2.5 2.8L14.5 10" fill="none" stroke="' + (color || C.k) + '" stroke-width="1.2"/>'
      : '<path d="M12 2.5v10M9.5 10l2.5 2.8L14.5 10" fill="none" stroke="' + (color || C.k) + '" stroke-width="1.2"/>';
  }

  function sortAZ16(desc) {
    return (
      text(1, 7, desc ? "Z" : "A", 7, desc ? C.k : C.b) + text(1, 15, desc ? "A" : "Z", 7, desc ? C.b : C.k) + sortArrow16(desc)
    );
  }

  function cylinder32(x, y, w, h, fill) {
    const rx = w / 2;
    const ry = 3;
    return (
      '<path d="M' + x + " " + (y + ry) + "v" + (h - ry * 2) + "a" + rx + " " + ry + " 0 0 0 " + w + " 0v" + -(h - ry * 2) +
      '" fill="' + (fill || "#fff") + '" stroke="' + C.k + '"/>' +
      '<ellipse cx="' + (x + rx) + '" cy="' + (y + ry) + '" rx="' + rx + '" ry="' + ry + '" fill="' + (fill || "#fff") + '" stroke="' + C.k + '"/>'
    );
  }

  function globe(cx, cy, r, color) {
    const c = color || C.b;
    return (
      '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '" fill="#fff" stroke="' + c + '"/>' +
      '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + r * 0.45 + '" ry="' + r + '" fill="none" stroke="' + c + '"/>' +
      '<path d="M' + (cx - r) + " " + cy + "H" + (cx + r) + '" stroke="' + c + '"/>'
    );
  }

  function clock(cx, cy, r) {
    return (
      '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '" fill="#fff" stroke="' + C.b + '" stroke-width="1.2"/>' +
      '<path d="M' + cx + " " + (cy - r * 0.6) + "V" + cy + "h" + r * 0.5 + '" fill="none" stroke="' + C.b + '" stroke-width="1.2"/>'
    );
  }

  function calendar16() {
    return (
      '<rect x="1.5" y="3.5" width="13" height="11" fill="#fff" stroke="' + C.k + '"/><rect x="1.5" y="3.5" width="13" height="3" fill="' + C.r + '"/>' +
      '<path d="M4.5 1.5v3M11.5 1.5v3" stroke="' + C.k + '"/><path d="M4 9h2M7 9h2M10 9h2M4 12h2M7 12h2" stroke="' + C.kl + '"/>'
    );
  }

  function pencil(x, y, s, color) {
    return '<path d="M' + x + " " + (y + s) + "l" + s * 0.2 + " " + -s * 0.55 + "l" + s * 0.6 + " " + -s * 0.6 + "l" + s * 0.35 + " " + s * 0.35 + "l" + -s * 0.6 + " " + s * 0.6 + 'z" fill="' + (color || C.y) + '" stroke="' + C.k + '" stroke-width="0.8"/>';
  }

  function arrowRight(x1, y, x2, color, sw) {
    return (
      '<path d="M' + x1 + " " + y + "H" + x2 + "M" + (x2 - 3) + " " + (y - 3) + "L" + x2 + " " + y + "L" + (x2 - 3) + " " + (y + 3) +
      '" fill="none" stroke="' + (color || C.g) + '" stroke-width="' + (sw || 1.8) + '" stroke-linecap="round" stroke-linejoin="round"/>'
    );
  }

  function bars(x, y, color) {
    return (
      '<rect x="' + x + '" y="' + (y + 6) + '" width="3" height="6" fill="' + (color || C.b) + '"/>' +
      '<rect x="' + (x + 4) + '" y="' + (y + 2) + '" width="3" height="10" fill="' + (color || C.b) + '"/>' +
      '<rect x="' + (x + 8) + '" y="' + (y + 4) + '" width="3" height="8" fill="' + (color || C.b) + '"/>'
    );
  }

  function fx(x, y, size) {
    return text(x, y, "<tspan font-style=\"italic\">fx</tspan>", size, C.k, 600);
  }

  /* ---------------- Symbole ---------------- */

  const BIG = {
    // Excel › Daten
    dataGet: () => cylinder32(3.5, 3.5, 15, 20) + table32({ x: 12.5, y: 13.5, w: 16, h: 14, cols: 2, rows: 3 }),
    refreshAll: () => doc32() + refresh(16.5, 19, 6),
    stocks: () =>
      '<path d="M4 12.5L16 5l12 7.5z" fill="#fff" stroke="' + C.k + '" stroke-linejoin="round"/>' +
      '<path d="M7.5 14v10M12.5 14v10M19.5 14v10M24.5 14v10" stroke="' + C.k + '" stroke-width="2"/><path d="M4 26.5h24" stroke="' + C.k + '" stroke-width="2"/>',
    currencies: () =>
      '<rect x="3.5" y="9.5" width="25" height="15" rx="1" fill="#fff" stroke="' + C.k + '"/><circle cx="16" cy="17" r="4.5" fill="' + C.gl + '" stroke="' + C.g + '"/>' +
      text(13.8, 20, "€", 7, C.gd),
    sort: () =>
      text(3, 13, "A", 10, C.b) + text(12, 13, "Z", 10, C.k) + text(3, 27, "Z", 10, C.k) + text(12, 27, "A", 10, C.b) +
      '<path d="M25 5v20M21.5 21.5L25 25.5l3.5-4" fill="none" stroke="' + C.k + '" stroke-width="1.4"/>',
    filter: () => funnel32(),
    textToColumns: () =>
      table32({ x: 3.5, y: 6.5, w: 11, h: 20, cols: 1, rows: 4 }) + table32({ x: 18.5, y: 6.5, w: 10, h: 20, cols: 2, rows: 4 }) + arrowRight(13, 16.5, 20, C.b),
    cleanData: () =>
      table32({ x: 3.5, y: 9.5, w: 20, h: 18, cols: 3, rows: 4 }) +
      '<path d="M25 3l1.2 3 3 1.2-3 1.2L25 11.5l-1.2-3-3-1.2 3-1.2z" fill="' + C.p + '"/>' +
      '<path d="M20 11l.8 1.8 1.8.8-1.8.8L20 16.3l-.8-1.9-1.8-.8 1.8-.8z" fill="' + C.b + '"/>',
    whatIf: () =>
      table32({ x: 3.5, y: 6.5, w: 18, h: 16, cols: 2, rows: 3 }) + '<circle cx="22" cy="21" r="7" fill="#fff" stroke="' + C.b + '" stroke-width="1.4"/>' + text(19.4, 25, "?", 10, C.b),
    forecastSheet: () =>
      '<rect x="3.5" y="4.5" width="25" height="23" fill="#fff" stroke="' + C.k + '"/>' +
      '<path d="M6 22l5-6 4 3 5-7" fill="none" stroke="' + C.b + '" stroke-width="1.8"/><path d="M20 12l6-5" fill="none" stroke="' + C.o + '" stroke-width="1.8" stroke-dasharray="2 1.5"/>',
    outline: () =>
      '<path d="M5.5 4.5v23M5.5 8.5h4M5.5 16.5h4M5.5 24.5h4" stroke="' + C.k + '"/><rect x="11.5" y="5.5" width="17" height="6" fill="#fff" stroke="' + C.k + '"/>' +
      '<rect x="11.5" y="13.5" width="17" height="6" fill="' + C.bl + '" stroke="' + C.k + '"/><rect x="11.5" y="21.5" width="17" height="6" fill="#fff" stroke="' + C.k + '"/>',

    // Power Query › Start
    closeLoad: () =>
      table32({ x: 11.5, y: 3.5, w: 17, h: 16, cols: 2, rows: 3 }) +
      '<rect x="3.5" y="13.5" width="14" height="15" rx="1" fill="' + C.p + '"/><rect x="6.5" y="13.5" width="8" height="5" fill="#fff"/><rect x="6" y="21.5" width="9" height="6" fill="#E8E2F7"/>' +
      '<path d="M22 22v6M19 25l3 3 3-3" fill="none" stroke="' + C.g + '" stroke-width="1.8"/>',
    refreshPreview: () => doc32() + refresh(16.5, 19, 6, C.b),
    chooseColumns: () => table32({ selCol: 1, selFill: "#BFD9F2" }) + check(16, 18, 3, C.g),
    removeColumns: () => table32({ selCol: 1, selFill: "#F4D3D3" }) + xMark(16, 18, 4.5),
    keepRows: () => table32({ selRow: 2, selFill: C.gl }) + check(16, 16.5, 3),
    removeRows: () => table32({ selRow: 2, selFill: "#F4D3D3" }) + xMark(16, 16.5, 4),
    splitColumn: () =>
      '<rect x="4.5" y="4.5" width="10" height="23" fill="#fff" stroke="' + C.k + '"/><rect x="4.5" y="4.5" width="10" height="5" fill="' + C.g + '"/>' +
      '<rect x="18.5" y="4.5" width="10" height="23" fill="#fff" stroke="' + C.k + '"/><rect x="18.5" y="4.5" width="10" height="5" fill="' + C.g + '"/>' +
      '<path d="M16.5 2v28" stroke="' + C.r + '" stroke-width="1.6" stroke-dasharray="2.5 2"/>',
    groupBy: () =>
      table32({ x: 3.5, y: 4.5, w: 12, h: 12, cols: 1, rows: 3 }) + table32({ x: 3.5, y: 17.5, w: 12, h: 10, cols: 1, rows: 2 }) +
      '<path d="M17 10.5h3v12h-3M20 16.5h3" fill="none" stroke="' + C.k + '"/>' + '<rect x="23.5" y="12.5" width="6" height="8" fill="' + C.gl + '" stroke="' + C.g + '"/>',
    manageParameters: () => doc32() + '<path d="M11 14h10M11 18h10M11 22h7" stroke="' + C.kl + '" stroke-width="1.4"/>' + gear(21, 23, 4.5, C.b),
    dataSourceSettings: () => doc32() + gear(20, 22, 5, C.o),
    // Transformieren
    transpose: () => table32({ cols: 3, rows: 3 }) + '<path d="M9 22a9 9 0 0 1 14-10" fill="none" stroke="' + C.b + '" stroke-width="1.6"/>',
    reverseRows: () => table32() + '<path d="M16 11v12M13 20l3 3.5 3-3.5M13 14l3-3.5 3 3.5" fill="none" stroke="' + C.b + '" stroke-width="1.6"/>',
    countRows: () => table32({ cols: 2 }) + text(17, 25, "123", 8, C.b),
    detectType: () => table32({ cols: 2 }) + text(9, 22, "?", 12, C.b),
    pivot: () => table32() + '<path d="M9 23V14h12" fill="none" stroke="' + C.b + '" stroke-width="1.8"/><path d="M18 11l3 3-3 3" fill="none" stroke="' + C.b + '" stroke-width="1.8"/>',
    unpivot: () => table32() + '<path d="M9 14h11v9" fill="none" stroke="' + C.b + '" stroke-width="1.8"/><path d="M17 20l3 3 3-3" fill="none" stroke="' + C.b + '" stroke-width="1.8"/>',
    format: () => text(3, 20, "A", 14, C.k) + text(14, 26, "a", 12, C.b) + pencil(20, 3, 9),
    mergeColumns: () =>
      '<rect x="3.5" y="5.5" width="8" height="20" fill="#fff" stroke="' + C.k + '"/><rect x="12.5" y="5.5" width="8" height="20" fill="#fff" stroke="' + C.k + '"/>' +
      '<rect x="3.5" y="5.5" width="17" height="4" fill="' + C.g + '"/>' + arrowRight(21, 15.5, 29, C.b),
    statistics: () => text(4, 24, "Σ", 20, C.k, 400),
    standard: () => text(3, 16, "+", 12, C.b) + text(15, 16, "−", 12, C.b) + text(3, 29, "×", 12, C.b) + text(16, 29, "÷", 12, C.b),
    scientific: () => text(4, 23, "10", 13, C.k, 600) + text(19, 14, "2", 9, C.b),
    date: () => calendarBig(),
    time: () => '<circle cx="16" cy="16" r="12" fill="#fff" stroke="' + C.b + '" stroke-width="1.4"/><path d="M16 8v8h6" fill="none" stroke="' + C.b + '" stroke-width="1.8"/>',
    duration: () => '<path d="M9 4.5h14M9 27.5h14M10 5c0 8 12 8 12 11S10 19 10 27M22 5c0 8-12 8-12 11s12 3 12 11" fill="none" stroke="' + C.k + '" stroke-width="1.3"/><path d="M12 25h8l-4-5z" fill="' + C.y + '"/>',
    expand: () => table32({ cols: 2 }) + '<path d="M18 17h8M22 13v8" stroke="' + C.g + '" stroke-width="2"/>',
    aggregate: () => table32({ cols: 2 }) + text(15, 26, "Σ", 11, C.g, 400),
    // Spalte hinzufügen
    columnFromExamples: () => table32({ selCol: 2 }) + pencil(18, 13, 11),
    customColumn: () => table32({ selCol: 2 }) + fx(17, 24, 10),
    invokeFunction: () => doc32() + fx(10, 23, 11),
    conditionalColumn: () =>
      table32({ selCol: 2 }) + '<path d="M22 12l5 5-5 5-5-5z" fill="' + C.y + '" stroke="' + C.k + '" stroke-width="0.8"/>',
    indexColumn: () => table32({ selCol: 0 }) + text(5.5, 20, "1", 6, C.b) + text(5.5, 25.5, "2", 6, C.b),
    duplicateColumn: () =>
      '<rect x="4.5" y="4.5" width="11" height="20" fill="#fff" stroke="' + C.k + '"/><rect x="4.5" y="4.5" width="11" height="5" fill="' + C.g + '"/>' +
      '<rect x="16.5" y="8.5" width="11" height="20" fill="' + C.gl + '" stroke="' + C.k + '"/><rect x="16.5" y="8.5" width="11" height="5" fill="' + C.g + '"/>',
    // Ansicht
    queryDependencies: () =>
      '<rect x="3.5" y="3.5" width="10" height="7" fill="#fff" stroke="' + C.k + '"/><rect x="18.5" y="12.5" width="10" height="7" fill="' + C.gl + '" stroke="' + C.g + '"/>' +
      '<rect x="3.5" y="21.5" width="10" height="7" fill="#fff" stroke="' + C.k + '"/><path d="M13.5 7h3v12.5h2M13.5 25h3V16" fill="none" stroke="' + C.k + '"/>',
    advancedEditorBig: () => doc32() + text(10, 23, "M", 11, C.g),
    goToColumn: () => table32({ selCol: 1 }) + arrowRight(4, 30, 14, C.b, 1.6),
    useFirstRowBig: () =>
      table32({ hdr: "#fff" }) + '<rect x="3.5" y="6.5" width="25" height="5" fill="' + C.gl + '" stroke="' + C.g + '"/>' +
      '<path d="M16 24v-9M12.5 18.5L16 15l3.5 3.5" fill="none" stroke="' + C.g + '" stroke-width="1.8"/>',
    queriesConnectionsBig: () =>
      '<rect x="3.5" y="4.5" width="25" height="23" fill="#fff" stroke="' + C.k + '"/><rect x="17.5" y="4.5" width="11" height="23" fill="' + C.gl + '" stroke="' + C.k + '"/>' +
      '<path d="M6.5 10h8M6.5 15h8M6.5 20h8" stroke="' + C.kl + '" stroke-width="1.4"/>',
    emptyBig: () => table32(),
  };

  function calendarBig() {
    return (
      '<rect x="3.5" y="6.5" width="25" height="22" fill="#fff" stroke="' + C.k + '"/><rect x="3.5" y="6.5" width="25" height="6" fill="' + C.r + '"/>' +
      '<path d="M9.5 3.5v5M22.5 3.5v5" stroke="' + C.k + '" stroke-width="1.4"/>' +
      '<path d="M7 17h3M12 17h3M17 17h3M22 17h3M7 22h3M12 22h3M17 22h3" stroke="' + C.kl + '" stroke-width="2"/>'
    );
  }

  const SMALL = {
    // Excel › Daten
    fromTextCsv: () => doc16() + '<path d="M5 7h5M5 9.5h5M5 12h3" stroke="' + C.g + '"/>',
    fromWeb: () => globe(8, 8, 6),
    fromTable: () => table16({ hdr: C.b }) + '<rect x="10" y="9" width="5.5" height="5.5" fill="' + C.g + '"/>' + '<path d="M11.5 11.8h2.5M12.8 10.5v2.6" stroke="#fff"/>',
    fromPicture: () => '<rect x="1.5" y="2.5" width="13" height="11" fill="#fff" stroke="' + C.k + '"/><circle cx="5" cy="6" r="1.4" fill="' + C.y + '"/><path d="M2 13l4-4 3 3 2-2 3.5 3" fill="' + C.gl + '" stroke="' + C.g + '"/>',
    recentSources: () => doc16() + clock(10.5, 11, 3.8),
    existingConnections: () => '<rect x="1.5" y="4.5" width="9" height="10" fill="#fff" stroke="' + C.k + '"/><rect x="5.5" y="1.5" width="9" height="10" fill="#fff" stroke="' + C.k + '"/><path d="M7.5 5h5M7.5 7.5h5" stroke="' + C.b + '"/>',
    queriesConnections: () => '<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="' + C.k + '"/><rect x="9.5" y="1.5" width="5" height="13" fill="' + C.gl + '" stroke="' + C.k + '"/><path d="M3.5 5h4M3.5 8h4M3.5 11h4" stroke="' + C.kl + '"/>',
    properties: () => '<rect x="2.5" y="1.5" width="11" height="13" fill="#fff" stroke="' + C.k + '"/><path d="M4.5 5h1M7 5h4.5M4.5 8h1M7 8h4.5M4.5 11h1M7 11h4.5" stroke="' + C.k + '"/>',
    editLinks: () => '<path d="M6.5 9.5l3-3" stroke="' + C.k + '" stroke-width="1.3"/><rect x="1.5" y="7" width="7" height="4" rx="2" transform="rotate(-45 5 9)" fill="none" stroke="' + C.b + '" stroke-width="1.3"/><rect x="7.5" y="5" width="7" height="4" rx="2" transform="rotate(-45 11 7)" fill="none" stroke="' + C.b + '" stroke-width="1.3"/>',
    sortAZ: () => sortAZ16(false),
    sortZA: () => sortAZ16(true),
    clearFilter: () => funnel16() + xMark(12, 12, 2.4, 1.6),
    reapply: () => funnel16() + refresh(12, 12, 2.8, C.g, 1.3),
    advancedFilter: () => funnel16() + pencil(9, 8, 7),
    flashFill: () => table16() + '<path d="M11 6l-3 5h2.5l-1 4 4-6H11l1.5-3z" fill="' + C.y + '" stroke="' + C.k + '" stroke-width="0.6"/>',
    removeDuplicates: () => table16({ hdr: C.b }) + xMark(11.5, 11, 2.6, 1.6),
    dataValidation: () => table16({ hdr: C.b }) + check(10.5, 11, 2.2, C.g) + xMark(4.5, 10.5, 1.6, 1.3),
    consolidate: () => '<rect x="1.5" y="1.5" width="5" height="5" fill="#fff" stroke="' + C.k + '"/><rect x="1.5" y="9.5" width="5" height="5" fill="#fff" stroke="' + C.k + '"/><rect x="10.5" y="5.5" width="4" height="5" fill="' + C.gl + '" stroke="' + C.g + '"/><path d="M6.5 4h2v8h-2M8.5 8h2" fill="none" stroke="' + C.k + '"/>',
    relationships: () => '<rect x="1.5" y="1.5" width="6" height="5" fill="#fff" stroke="' + C.k + '"/><rect x="8.5" y="9.5" width="6" height="5" fill="#fff" stroke="' + C.k + '"/><path d="M4.5 6.5v5.5h4" fill="none" stroke="' + C.b + '" stroke-width="1.2"/>',
    manageDataModel: () => cylinderSmall() + '<rect x="9.5" y="8.5" width="6" height="6" fill="' + C.g + '"/>',
    dataAnalysis: () => '<rect x="1.5" y="2.5" width="13" height="11" fill="#fff" stroke="' + C.k + '"/>' + bars(3, 1, C.b),
    // Power Query › Start
    properties2: () => '<rect x="2.5" y="1.5" width="11" height="13" fill="#fff" stroke="' + C.k + '"/><path d="M4.5 5h7M4.5 8h7M4.5 11h4" stroke="' + C.k + '"/>',
    advancedEditor: () => '<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="' + C.k + '"/><path d="M4 5h5M5.5 7.5h6M5.5 10h4M4 12.5h3" stroke="' + C.g + '"/>',
    manage: () => '<rect x="1.5" y="2.5" width="13" height="11" fill="#fff" stroke="' + C.k + '"/><path d="M3.5 5.5h9M3.5 8h9M3.5 10.5h9" stroke="' + C.kl + '"/>',
    sortAsc: () => sortAZ16(false),
    sortDesc: () => sortAZ16(true),
    dataType: () => '<rect x="1.5" y="2.5" width="13" height="11" fill="#fff" stroke="' + C.k + '"/>' + text(2.8, 10.8, "1", 7, C.b) + text(7.2, 10.8, "2", 7, C.g) + '<path d="M3 12.5h10" stroke="' + C.kl + '"/>',
    useFirstRow: () => table16({ hdr: "#fff" }) + '<rect x="1.5" y="2.5" width="13" height="3" fill="' + C.gl + '" stroke="' + C.g + '"/><path d="M8 13V7M5.5 9.5L8 7l2.5 2.5" fill="none" stroke="' + C.g + '" stroke-width="1.2"/>',
    replaceValues: () => text(0.5, 8, "1", 8, C.b) + text(8.5, 15.5, "2", 8, C.g) + '<path d="M7 3.5h4.5v4M9.5 6l2 2 2-2" fill="none" stroke="' + C.k + '"/>',
    mergeQueries: () => table16({ w: 7, x: 1.5 }) + table16({ w: 7, x: 7.5, hdr: C.b }),
    appendQueries: () => '<rect x="2.5" y="1.5" width="11" height="6" fill="#fff" stroke="' + C.k + '"/><rect x="2.5" y="1.5" width="11" height="2" fill="' + C.g + '"/><rect x="2.5" y="8.5" width="11" height="6" fill="' + C.bl + '" stroke="' + C.k + '"/>',
    combineFiles: () => doc16() + '<path d="M5.5 8.5h5M8 6v5" stroke="' + C.g + '" stroke-width="1.4"/>',
    newSource: () => doc16() + '<circle cx="11.5" cy="11.5" r="3.5" fill="' + C.y + '"/>' + plus(11.5, 11.5, 1.8, "#fff", 1.3),
    enterData: () => table16() + pencil(8, 7, 8),
    // Transformieren
    rename: () => '<rect x="1.5" y="4.5" width="13" height="7" fill="#fff" stroke="' + C.k + '"/>' + text(3, 10.5, "ab", 6, C.b) + '<path d="M10.5 3v10M9 3h3M9 13h3" stroke="' + C.k + '"/>',
    fillDown: () => table16() + '<path d="M8 6v7M5.5 10.5L8 13l2.5-2.5" fill="none" stroke="' + C.b + '" stroke-width="1.3"/>',
    move: () => table16() + '<path d="M3.5 9h9M10.5 7l2 2-2 2" fill="none" stroke="' + C.b + '" stroke-width="1.3"/>',
    toList: () => '<path d="M2 3.5h2M6 3.5h8M2 8h2M6 8h8M2 12.5h2M6 12.5h8" stroke="' + C.k + '" stroke-width="1.4"/>',
    extract: () => '<rect x="1.5" y="4.5" width="13" height="7" fill="#fff" stroke="' + C.k + '"/><rect x="5.5" y="4.5" width="5" height="7" fill="' + C.gl + '" stroke="' + C.g + '"/>',
    parse: () => doc16() + text(4.2, 12.5, "{ }", 6, C.b),
    trig: () => '<path d="M1 8c2-6 4-6 7 0s5 6 7 0" fill="none" stroke="' + C.b + '" stroke-width="1.4"/>',
    rounding: () => text(0.5, 12, ".00", 8, C.k),
    information: () => '<circle cx="8" cy="8" r="6.5" fill="#fff" stroke="' + C.b + '"/><path d="M8 7v5" stroke="' + C.b + '" stroke-width="1.6"/><circle cx="8" cy="4.6" r="1" fill="' + C.b + '"/>',
    extractValues: () => table16() + text(9, 14, "…", 7, C.b),
    // Allgemein
    table: () => table16(),
    query: () => table16({ hdr: C.g }),
    calendar: () => calendar16(),
    gearSmall: () => gear(8, 8, 6, C.k),
    close: () => '<path d="M4 4l8 8M12 4l-8 8" stroke="' + C.k + '" stroke-width="1.2"/>',
    fxSmall: () => fx(1, 12, 11),
    info: () => '<circle cx="8" cy="8" r="7" fill="' + C.b + '"/><path d="M8 7v5" stroke="#fff" stroke-width="1.6"/><circle cx="8" cy="4.5" r="1" fill="#fff"/>',
    warning: () => '<path d="M8 1.5l7 13H1z" fill="' + C.y + '" stroke="#9A6A00" stroke-linejoin="round"/><path d="M8 6v4.5" stroke="#1F1F1F" stroke-width="1.5"/><circle cx="8" cy="12.3" r="0.9" fill="#1F1F1F"/>',
    save: () => '<rect x="2" y="2" width="12" height="12" rx="1" fill="' + C.p + '"/><rect x="4.5" y="2" width="7" height="4" fill="#fff"/><rect x="4" y="8.5" width="8" height="5.5" fill="#E9E3F7"/>',
    undo: () => '<path d="M5 6h5.5a3.5 3.5 0 0 1 0 7H6" fill="none" stroke="' + C.k + '" stroke-width="1.3"/><path d="M6.5 3L3.5 6l3 3" fill="none" stroke="' + C.k + '" stroke-width="1.3"/>',
    redo: () => '<path d="M11 6H5.5a3.5 3.5 0 0 0 0 7H10" fill="none" stroke="' + C.k + '" stroke-width="1.3"/><path d="M9.5 3l3 3-3 3" fill="none" stroke="' + C.k + '" stroke-width="1.3"/>',
    search: () => '<circle cx="6.5" cy="6.5" r="4.5" fill="none" stroke="' + C.k + '" stroke-width="1.2"/><path d="M10 10l4.5 4.5" stroke="' + C.k + '" stroke-width="1.3"/>',
    comment: () => '<path d="M2.5 3.5h11v7h-6l-3 3v-3h-2z" fill="#fff" stroke="' + C.k + '"/>',
    share: () => '<path d="M3.5 6.5v7h9v-7" fill="none" stroke="#fff"/><path d="M8 10V2M5.5 4.5L8 2l2.5 2.5" fill="none" stroke="#fff" stroke-width="1.2"/>',
    excelApp: () => '<rect x="4" y="1.5" width="10.5" height="13" rx="1" fill="#fff" stroke="' + C.g + '"/><rect x="1" y="4" width="8" height="8" rx="1" fill="' + C.g + '"/><path d="M3 6l4 4M7 6l-4 4" stroke="#fff" stroke-width="1.3"/><path d="M10 5h3M10 8h3M10 11h3" stroke="' + C.g + '"/>',
    smiley: () => '<circle cx="8" cy="8" r="7" fill="' + C.y + '"/><circle cx="5.5" cy="6.5" r="1" fill="#5A3E00"/><circle cx="10.5" cy="6.5" r="1" fill="#5A3E00"/><path d="M4.5 9.5a3.8 3.8 0 0 0 7 0" fill="none" stroke="#5A3E00" stroke-width="1.1"/>',
    chevronLeft: () => '<path d="M10 3L5 8l5 5" fill="none" stroke="' + C.k + '" stroke-width="1.5"/>',
    chevronRight: () => '<path d="M6 3l5 5-5 5" fill="none" stroke="' + C.k + '" stroke-width="1.5"/>',
    help: () => '<circle cx="8" cy="8" r="7" fill="' + C.b + '"/>' + text(5.3, 12, "?", 10, "#fff"),
    filterArrow: () => '<path d="M4 6h8l-4 5z" fill="' + C.k + '"/>',
    funnelActive: () => funnel16(C.bl),
  };

  function cylinderSmall() {
    return '<path d="M1.5 3.5v8a4.5 2 0 0 0 9 0v-8" fill="#fff" stroke="' + C.k + '"/><ellipse cx="6" cy="3.5" rx="4.5" ry="2" fill="#fff" stroke="' + C.k + '"/>';
  }

  // Typ-Symbole in den Spaltenköpfen des Power Query-Editors
  const TYPE_GLYPHS = {
    Text: () => svg(16, text(0, 11.5, "A", 7.5, C.k) + text(4.3, 11.5, "B", 7.5, C.k) + text(9, 11.5, "C", 7.5, C.k)),
    Zahl: () => svg(16, text(0.2, 11.5, "123", 7.2, C.k)),
    Dezimalzahl: () => svg(16, text(0.6, 11.5, "1.2", 7.8, C.k)),
    Datum: () => svg(16, calendar16()),
    Beliebig: () => svg(16, text(0, 11.5, "ABC", 5.4, C.k) + text(2, 15.5, "123", 5, C.k)),
  };

  function get(name, size) {
    const set = size === 32 ? BIG : SMALL;
    const fn = set[name] || (size === 32 ? BIG.emptyBig : SMALL.table);
    return svg(size === 32 ? 32 : 16, fn());
  }

  function typeGlyph(type) {
    return (TYPE_GLYPHS[type] || TYPE_GLYPHS.Beliebig)();
  }

  window.ExcelFloIcons = { get, typeGlyph, has: (name, size) => !!(size === 32 ? BIG : SMALL)[name] };
})();
