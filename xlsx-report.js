/* ── Excel report builder (shared: browser + node) ─────────────────────
   buildCountWorkbook(ExcelJS, sections, meta) → ExcelJS.Workbook
   sections: [{ title, icon, color, rows:[{name, code, parts:[n], boxes, coeff, kg, extra}] }]
   meta:     { date:'dd.mm.yyyy', time:'HH:MM', branch:'118' }
   Layout: two item columns per page (A–F right, H–M left), each item framed
   with a thick border and a large scannable EAN-13 barcode under its name.
*/
(function (root) {
  const GREEN = 'FF1B5E20', LIGHT = 'FFE8F5E9', GREY = 'FF9E9E9E', BORDER = 'FFBDBDBD';
  const thin = { style: 'thin', color: { argb: BORDER } };
  const thick = { style: 'thick', color: { argb: 'FF000000' } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const r1 = n => Math.round(n * 10) / 10;
  const colL = c => String.fromCharCode(64 + c);

  /* ── EAN-13 → 1-bit PNG (base64), no canvas needed ── */
  const L = ['0001101', '0011001', '0010011', '0111101', '0100011', '0110001', '0101111', '0111011', '0110111', '0001011'];
  const R = L.map(p => p.replace(/./g, b => b === '0' ? '1' : '0'));
  const G = R.map(p => p.split('').reverse().join(''));
  const PARITY = ['LLLLLL', 'LLGLGG', 'LLGGLG', 'LLGGGL', 'LGLLGG', 'LGGLLG', 'LGGGLL', 'LGLGLG', 'LGLGGL', 'LGGLGL'];
  function ean13Bits(code) {
    const d = code.split('').map(Number);
    let s = '101';
    for (let i = 1; i <= 6; i++) s += (PARITY[d[0]][i - 1] === 'L' ? L : G)[d[i]];
    s += '01010';
    for (let i = 7; i <= 12; i++) s += R[d[i]];
    return s + '101';
  }
  const CRC = (() => {
    const t = [];
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  const crc32 = bytes => { let c = 0xFFFFFFFF; for (const b of bytes) c = CRC[(c ^ b) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const u32 = n => [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];
  const chunk = (type, data) => {
    const td = [...type].map(ch => ch.charCodeAt(0)).concat(data);
    return u32(data.length).concat(td, u32(crc32(td)));
  };
  const MODULE = 6, QUIET = 10, BC_H = 150;           // px; shown in Excel at half size → crisp print
  function barcodePngB64(code) {
    if (!/^\d{13}$/.test(code || '')) return null;
    const bits = '0'.repeat(QUIET) + ean13Bits(code) + '0'.repeat(QUIET);
    const w = bits.length * MODULE, rowLen = 1 + Math.ceil(w / 8);
    const line = new Array(rowLen).fill(0);              // filter byte 0, then 1 = white
    for (let x = 0; x < w; x++) if (bits[Math.floor(x / MODULE)] === '0') line[1 + (x >> 3)] |= 0x80 >> (x & 7);
    for (let x = w; x < (rowLen - 1) * 8; x++) line[1 + (x >> 3)] |= 0x80 >> (x & 7);
    const raw = [];
    for (let y = 0; y < BC_H; y++) raw.push(...line);
    const z = [0x78, 0x01];                              // zlib, stored blocks
    for (let i = 0; i < raw.length; i += 65535) {
      const blk = raw.slice(i, i + 65535), last = i + 65535 >= raw.length ? 1 : 0;
      z.push(last, blk.length & 255, blk.length >> 8, ~blk.length & 255, (~blk.length >> 8) & 255, ...blk);
    }
    let a = 1, b = 0;
    for (const v of raw) { a = (a + v) % 65521; b = (b + a) % 65521; }
    z.push(...u32(((b << 16) | a) >>> 0));
    const png = [137, 80, 78, 71, 13, 10, 26, 10]
      .concat(chunk('IHDR', u32(w).concat(u32(BC_H), [1, 0, 0, 0, 0])), chunk('IDAT', z), chunk('IEND', []));
    if (typeof Buffer !== 'undefined') return Buffer.from(png).toString('base64');
    let s = '';
    for (let i = 0; i < png.length; i += 0x8000) s += String.fromCharCode.apply(null, png.slice(i, i + 0x8000));
    return btoa(s);
  }

  /* ── layout ── */
  const WIDTHS = [4, 20, 8.5, 7, 7, 8.5];                  // # | name | parts | boxes | coeff | kg
  const OFF = [1, 8];                                    // first column of right / left item column
  const H_NAME = 36, H_BC = 62, H_DIG = 16, H_GAP = 6, H_SEC = 26;
  const PAGE_PT = 840;                                   // usable height (A4, fit to width ≈ 85%)

  function buildCountWorkbook(ExcelJS, sections, meta) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ספירת ירקות ופירות';
    wb.created = new Date();
    const ws = wb.addWorksheet('ספירה', {
      views: [{ rightToLeft: true, state: 'frozen', ySplit: 5 }],
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.35, right: 0.35, top: 0.4, bottom: 0.5, header: 0.2, footer: 0.25 },
        printTitlesRow: '5:5',
      },
    });
    ws.headerFooter.oddFooter = '&Cעמוד &P מתוך &N';
    ws.columns = [...WIDTHS, 2, ...WIDTHS].map(width => ({ width }));
    const LAST = 13;

    let counted = 0, totalKg = 0;
    sections.forEach(s => s.rows.forEach(r => { if (r.boxes > 0 || (r.extra && r.kg > 0)) { counted++; totalKg += r.kg || 0; } }));

    const banner = (row, value, font, height, bg) => {
      ws.mergeCells(row, 1, row, LAST);
      const c = ws.getCell(row, 1);
      c.value = value; c.font = font; c.alignment = { horizontal: 'center', vertical: 'middle' };
      if (bg) c.fill = fill(bg);
      ws.getRow(row).height = height;
    };
    banner(1, `ספירת ירקות ופירות — סניף ${meta.branch}`, { bold: true, size: 20, color: { argb: 'FFFFFFFF' } }, 34, 'FF2E7D32');
    banner(2, `תאריך: ${meta.date}    שעה: ${meta.time}`, { size: 13, color: { argb: GREEN } }, 20);
    banner(3, `פריטים שנספרו: ${counted}    |    סה"כ משקל: ${r1(totalKg)} ק"ג`, { bold: true, size: 13 }, 20, LIGHT);
    ws.getRow(4).height = 6;

    const HEAD = ['#', 'מוצר', 'ארגזים (לפי מיקום)', 'סה"כ ארגזים', 'מקדם ק"ג', 'סה"כ ק"ג'];
    const head = ws.getRow(5);
    head.height = 30;
    OFF.forEach(o => HEAD.forEach((h, i) => {
      const c = head.getCell(o + i);
      c.value = h;
      c.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
      c.fill = fill('FF388E3C');
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = box;
    }));

    let rowNo = 6, used = 34 + 20 + 20 + 6 + 30;
    const ensure = h => {                                // manual page break so no item is split
      if (used + h > PAGE_PT) { ws.getRow(rowNo - 1).addPageBreak(); used = 30; }
      used += h;
    };

    // thick frame around a rectangular block; inner lines thin
    const frame = (r0, r1_, c0, c1) => {
      for (let r = r0; r <= r1_; r++) for (let c = c0; c <= c1; c++) {
        const cell = ws.getCell(r, c);
        cell.border = {
          top: r === r0 ? thick : thin, bottom: r === r1_ ? thick : thin,
          right: c === c0 ? thick : thin, left: c === c1 ? thick : thin,
        };
      }
    };

    function item(r, o, it, num) {
      const has = it.boxes > 0 || (it.extra && it.kg > 0);
      const txt = { name: 'Arial', size: 12, color: { argb: has ? 'FF000000' : GREY } };
      const center = { horizontal: 'center', vertical: 'middle', wrapText: true };
      const set = (c, v, font, align) => { const cell = ws.getCell(r, o + c); cell.value = v; cell.font = font || txt; cell.alignment = align || center; if (has) cell.fill = fill(LIGHT); };
      const rows = it.extra ? 1 : 3;

      set(0, num);
      set(1, it.name, { name: 'Arial', size: 15, bold: true, color: { argb: has ? 'FF000000' : 'FF616161' } },
        { horizontal: 'right', vertical: 'middle', wrapText: true });
      set(2, !it.extra && it.parts.length > 1 ? it.parts.join(' + ') : null);
      set(3, !it.extra && it.boxes > 0 ? it.boxes : null);
      set(4, !it.extra && it.coeff ? it.coeff : null);

      if (rows > 1) ws.mergeCells(r, o + 5, r + rows - 1, o + 5);
      const kg = ws.getCell(r, o + 5);
      if (it.extra) kg.value = it.kg || null;
      else if (it.boxes > 0 && it.coeff) kg.value = { formula: `${colL(o + 3)}${r}*${colL(o + 4)}${r}`, result: r1(it.kg) };
      else if (it.boxes > 0 && (it.unit || /\(יח'?\)/.test(it.name))) kg.value = `${it.boxes} יח'`;   // per-unit item
      kg.numFmt = '0.0';
      kg.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      kg.font = has ? { name: 'Arial', size: 16, bold: true, color: { argb: 'FF000000' } } : txt;
      if (has) kg.fill = fill('FF00B050');

      if (!it.extra) {
        ws.mergeCells(r + 1, o, r + 1, o + 4);
        ws.mergeCells(r + 2, o, r + 2, o + 4);
        const b64 = barcodePngB64(it.code);
        if (b64) {
          const id = wb.addImage({ base64: b64, extension: 'png' });
          const w = 320, h = 72;                       // ≈ 0.46 scale of the 690×150 source
          ws.addImage(id, { tl: { col: o - 1 + 0.12, row: r + 0.06 }, ext: { width: w, height: h }, editAs: 'oneCell' });
          const d = ws.getCell(r + 2, o);
          d.value = `${it.code[0]}  ${it.code.slice(1, 7)}  ${it.code.slice(7)}`;
          d.font = { name: 'Arial', size: 12, bold: true };
          d.alignment = { horizontal: 'center', vertical: 'middle' };
        } else {
          const d = ws.getCell(r + 1, o);
          d.value = 'אין ברקוד במערכת';
          d.font = { name: 'Arial', size: 12, italic: true, color: { argb: GREY } };
          d.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      }
      frame(r, r + rows - 1, o, o + 5);
    }

    sections.forEach(sec => {
      if (!sec.rows.length) return;
      const extra = sec.rows[0].extra;
      const pairH = extra ? 28 + H_GAP : H_NAME + H_BC + H_DIG + H_GAP;
      ensure(H_SEC + pairH);                             // section title never alone at page bottom
      used -= pairH;
      banner(rowNo, `${sec.icon} ${sec.title}`, { bold: true, size: 15, color: { argb: GREEN } }, H_SEC, sec.color);
      rowNo++;

      const kgCells = [];
      for (let i = 0; i < sec.rows.length; i += 2) {
        ensure(pairH);
        if (extra) ws.getRow(rowNo).height = 28;
        else { ws.getRow(rowNo).height = H_NAME; ws.getRow(rowNo + 1).height = H_BC; ws.getRow(rowNo + 2).height = H_DIG; }
        [0, 1].forEach(k => {
          const it = sec.rows[i + k];
          if (!it) return;
          item(rowNo, OFF[k], it, i + k + 1);
          kgCells.push(`${colL(OFF[k] + 5)}${rowNo}`);
        });
        rowNo += extra ? 1 : 3;
        ws.getRow(rowNo).height = H_GAP;
        rowNo++;
      }

      ensure(24);
      const sub = ws.getRow(rowNo);
      sub.height = 24;
      sub.getCell(12).value = `סה"כ ${sec.title}:`;
      sub.getCell(13).value = { formula: `SUM(${kgCells.join(',')})`, result: r1(sec.rows.reduce((a, r) => a + (r.kg || 0), 0)) };
      sub.getCell(13).numFmt = '0.0';
      [12, 13].forEach(c => {
        sub.getCell(c).font = { bold: true, size: 13, color: { argb: GREEN } };
        sub.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' };
        sub.getCell(c).border = { top: { style: 'medium', color: { argb: GREEN } } };
      });
      rowNo++;
      ensure(14);
      ws.getRow(rowNo).height = 14;
      rowNo++;
    });
    return wb;
  }

  root.buildCountWorkbook = buildCountWorkbook;
  root.barcodePngB64 = barcodePngB64;
  if (typeof module !== 'undefined') module.exports = { buildCountWorkbook, barcodePngB64 };
})(typeof window !== 'undefined' ? window : globalThis);
