/* ── Printable manual count sheet for Avi (shared: browser + node) ────
   buildAviWorkbook(ExcelJS, vegetables, fruits, meta) → ExcelJS.Workbook
   meta: { branch:'118' }
   Four columns (RTL): name (Hebrew, English below) | scannable EAN-13 bars |
   barcode digits (one line) | empty box for the handwritten count.
   Large print (23pt names) for reading on paper and for photographing /
   scanning the filled pages later.
   The bars are drawn with CELL FILLS (one narrow column per EAN module), not
   as a floating picture: pictures are dropped when the file is printed from
   a mail/phone viewer or Google Sheets, cell fills print everywhere.
*/
(function (root) {
  const GREEN = 'FF2E7D32', GREY = 'FF616161', UNIT_BG = 'FFEEEEEE';
  const med = { style: 'medium', color: { argb: 'FF000000' } };
  const thick = { style: 'thick', color: { argb: 'FF000000' } };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const isUnit = it => !!it.unit || /\(יח'\)/.test(it.name);
  const byNameHe = (a, b) => a.name.localeCompare(b.name, 'he');

  const W_NAME = 40, W_CODE = 22, W_QTY = 18;
  const NAME_PT = 23, EN_PT = 12, CODE_PT = 13;
  const QUIET = 7, NBARS = 95 + QUIET * 2;         // bar columns: quiet | 95 modules | quiet
  const W_MOD = 0.25;                              // ≈3 px per module in Excel
  const C_NAME = 1, C_BAR0 = 2, C_BARN = C_BAR0 + NBARS - 1, C_CODE = C_BARN + 1, C_QTY = C_BARN + 2;
  const NCOL = C_QTY;
  const MIN_ROW_PT = 54;
  const BLACK = 'FF000000';

  /* ── EAN-13 → 95 modules ── */
  const L = ['0001101','0011001','0010011','0111101','0100011','0110001','0101111','0111011','0110111','0001011'];
  const G = ['0100111','0110011','0011011','0100001','0011101','0111001','0000101','0010001','0001001','0010111'];
  const R = L.map(s => s.replace(/./g, b => (b === '1' ? '0' : '1')));
  const PAR = ['LLLLLL','LLGLGG','LLGGLG','LLGGGL','LGLLGG','LGGLLG','LGGGLL','LGLGLG','LGLGGL','LGGLGL'];
  function ean13Modules(c) {
    if (!/^\d{13}$/.test(c)) return null;
    let sum = 0;
    for (let i = 0; i < 12; i++) sum += (+c[i]) * (i % 2 ? 3 : 1);
    if ((10 - sum % 10) % 10 !== +c[12]) return null;           // bad check digit → no bars
    const p = PAR[+c[0]];
    let m = '101';
    for (let i = 1; i <= 6; i++) m += (p[i - 1] === 'L' ? L : G)[+c[i]];
    m += '01010';
    for (let i = 7; i <= 12; i++) m += R[+c[i]];
    return m + '101';
  }

  /* ── minimal 1-bit grayscale PNG encoder (stored deflate) ── */
  const CRC = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (bytes) => { let c = 0xFFFFFFFF; for (const b of bytes) c = CRC[(c ^ b) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const u32 = v => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
  function chunk(type, data) {
    const td = [...type].map(ch => ch.charCodeAt(0)).concat(data);
    return u32(data.length).concat(td, u32(crc32(td)));
  }
  function pngBase64(width, height, isBlack) {          // isBlack(x,y) → bool
    const rowBytes = Math.ceil(width / 8), raw = [];
    for (let y = 0; y < height; y++) {
      raw.push(0);                                        // filter: none
      for (let bx = 0; bx < rowBytes; bx++) {
        let byte = 0;
        for (let k = 0; k < 8; k++) {
          const x = bx * 8 + k;
          const white = x >= width || !isBlack(x, y);
          byte |= (white ? 1 : 0) << (7 - k);             // 1 = white in grayscale
        }
        raw.push(byte);
      }
    }
    const z = [0x78, 0x01];
    for (let i = 0; i < raw.length; i += 65535) {
      const blk = raw.slice(i, i + 65535), last = i + 65535 >= raw.length ? 1 : 0;
      z.push(last, blk.length & 255, blk.length >>> 8, ~blk.length & 255, (~blk.length >>> 8) & 255);
      for (const b of blk) z.push(b);
    }
    let a = 1, b2 = 0;
    for (const v of raw) { a = (a + v) % 65521; b2 = (b2 + a) % 65521; }
    z.push(...u32(((b2 << 16) | a) >>> 0));
    const bytes = [137, 80, 78, 71, 13, 10, 26, 10]
      .concat(chunk('IHDR', u32(width).concat(u32(height), [1, 0, 0, 0, 0])))
      .concat(chunk('IDAT', z), chunk('IEND', []));
    if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
    let s = ''; for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode.apply(null, bytes.slice(i, i + 8192));
    return btoa(s);
  }
  function barcodePng(code) {
    const m = ean13Modules(code);
    if (!m) return null;
    const S = 3, QUIET = 9, H = 150, GUARD_H = 150, BAR_H = 136;
    const W = (95 + QUIET * 2) * S;
    const guard = i => i < 3 || (i >= 45 && i < 50) || i >= 92;
    return pngBase64(W, H, (x, y) => {
      const i = Math.floor(x / S) - QUIET;
      if (i < 0 || i >= 95 || m[i] !== '1') return false;
      return y < (guard(i) ? GUARD_H : BAR_H);
    });
  }

  function buildAviWorkbook(ExcelJS, vegetables, fruits, meta) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ספירת ירקות ופירות';
    wb.created = new Date();
    const ws = wb.addWorksheet('דף ספירה', {
      views: [{ rightToLeft: true }],
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.35, right: 0.35, top: 0.4, bottom: 0.5, header: 0.2, footer: 0.25 },
        printTitlesRow: '4:4',
      },
    });
    ws.headerFooter.oddFooter = '&C&14עמוד &P מתוך &N';
    ws.columns = [W_NAME, ...Array(NBARS).fill(W_MOD), W_CODE, W_QTY].map(width => ({ width }));

    const frame = (cell, top, bottom) => {
      cell.border = { top: top || med, bottom: bottom || med, left: thick, right: thick };
    };
    const banner = (row, value, font, height, bg) => {
      ws.mergeCells(row, 1, row, NCOL);
      const c = ws.getCell(row, 1);
      c.value = value; c.font = font;
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      if (bg) c.fill = fill(bg);
      ws.getRow(row).height = height;
    };

    banner(1, `ספירת ירקות ופירות — סניף ${meta.branch}`, { name: 'Arial', bold: true, size: 24, color: { argb: 'FFFFFFFF' } }, 42, GREEN);
    banner(2, 'תאריך: ____________          שם הסופר: אבי', { name: 'Arial', bold: true, size: 20 }, 40);
    banner(3, 'רשום בעמודה השמאלית כמה ארגזים יש.   ★ ברקע אפור = רשום כמה יחידות.',
      { name: 'Arial', size: 15, color: { argb: GREY } }, 30);

    const head = ws.getRow(4);
    head.height = 38;
    ws.mergeCells(4, C_BAR0, 4, C_BARN);
    [[C_NAME, 'שם המוצר'], [C_BAR0, 'סריקה'], [C_CODE, 'ברקוד'], [C_QTY, 'כמות']].forEach(([col, h]) => {
      const c = head.getCell(col);
      c.value = h;
      c.font = { name: 'Arial', bold: true, size: 22, color: { argb: 'FFFFFFFF' } };
      c.fill = fill('FF388E3C');
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      frame(c, thick, thick);
    });

    let r = 5;
    const section = title => {
      banner(r, title, { name: 'Arial', bold: true, size: 24, color: { argb: GREEN } }, 40, 'FFDCEDC8');
      frame(ws.getCell(r, 1), thick, thick);
      r++;
    };
    // name width in pt ≈ 7px per Excel width unit × 0.75, minus cell padding
    const nameLines = s => Math.max(1, Math.ceil(s.length * NAME_PT * 0.5 / (W_NAME * 5.25 - 12)));

    const itemRow = it => {
      const unit = isUnit(it);
      const code = String(it.barcodeNum || it.ean || '');
      const row = ws.getRow(r);
      const nameRich = [{ text: (unit ? '★ ' : '') + it.name, font: { name: 'Arial', bold: true, size: NAME_PT } }];
      if (it.en) nameRich.push({ text: '\n' + it.en, font: { name: 'Arial', size: EN_PT, color: { argb: GREY } } });
      row.height = Math.max(MIN_ROW_PT,
        nameLines((unit ? '★ ' : '') + it.name) * NAME_PT * 1.25 + (it.en ? EN_PT * 1.4 : 0) + 12);

      const a = row.getCell(C_NAME);
      a.value = { richText: nameRich };
      a.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true, readingOrder: 'rtl' };

      const mods = code ? ean13Modules(code) : null;
      const barCells = [];
      if (mods) {
        // sheet is RTL: higher column index = further left, so module 0 goes to the last bar column
        for (let k = 0; k < NBARS; k++) {
          const i = NBARS - 1 - k - QUIET;
          const cell = row.getCell(C_BAR0 + k);
          if (i >= 0 && i < 95 && mods[i] === '1') cell.fill = fill(BLACK);
          barCells.push(cell);
        }
      } else {
        ws.mergeCells(r, C_BAR0, r, C_BARN);
        const bars = row.getCell(C_BAR0);
        barCells.push(bars);
        bars.value = '—';
        bars.font = { name: 'Arial', size: 16, color: { argb: GREY } };
        bars.alignment = { horizontal: 'center', vertical: 'middle' };
      }

      const b = row.getCell(C_CODE);
      b.value = code || '—';
      b.font = { name: 'Arial', bold: true, size: CODE_PT, color: { argb: code ? 'FF000000' : GREY } };
      b.alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

      const c = row.getCell(C_QTY);
      c.value = { richText: [{ text: unit ? 'יחידות' : 'ארגזים', font: { name: 'Arial', size: 10, color: { argb: 'FFBDBDBD' } } }] };
      c.alignment = { horizontal: 'center', vertical: 'bottom' };
      if (unit) {
        a.fill = fill(UNIT_BG); b.fill = fill(UNIT_BG);
        if (!mods) barCells[0].fill = fill(UNIT_BG);      // bars keep a white background to scan
      }
      [a, b, c].forEach(x => frame(x));
      barCells.forEach(x => { x.border = { top: med, bottom: med }; });
      r++;
    };

    section('ירקות');
    vegetables.slice().sort(byNameHe).forEach(itemRow);
    section('פירות');
    fruits.slice().sort(byNameHe).forEach(itemRow);
    section('מוצרים נוספים (שלא ברשימה)');
    for (let i = 0; i < 6; i++) {
      const row = ws.getRow(r++);
      row.height = 56;
      ws.mergeCells(row.number, C_BAR0, row.number, C_BARN);
      [C_NAME, C_BAR0, C_CODE, C_QTY].forEach(k => frame(row.getCell(k)));
    }
    return wb;
  }

  root.buildAviWorkbook = buildAviWorkbook;
  root.aviBarcodePng = barcodePng;
  if (typeof module !== 'undefined' && module.exports) module.exports = { buildAviWorkbook, barcodePng };
})(typeof window !== 'undefined' ? window : globalThis);
