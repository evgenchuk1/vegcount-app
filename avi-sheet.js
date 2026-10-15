/* ── Printable manual count sheet for Avi (shared: browser + node) ────
   buildAviWorkbook(ExcelJS, vegetables, fruits, meta) → ExcelJS.Workbook
   meta: { branch:'118' }
   Three columns (RTL): name (Hebrew, English below) | barcode digits | empty box
   for the handwritten count. Large print (23pt names, 22pt barcode) for
   reading on paper and for photographing / scanning the filled pages later.
*/
(function (root) {
  const GREEN = 'FF2E7D32', GREY = 'FF616161', UNIT_BG = 'FFEEEEEE';
  const med = { style: 'medium', color: { argb: 'FF000000' } };
  const thick = { style: 'thick', color: { argb: 'FF000000' } };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const isUnit = it => !!it.unit || /\(יח'\)/.test(it.name);
  const byNameHe = (a, b) => a.name.localeCompare(b.name, 'he');

  const W_NAME = 40, W_CODE = 32, W_QTY = 26;           // A (right) | B | C (left)
  const NAME_PT = 23, EN_PT = 12, CODE_PT = 22;

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
    ws.columns = [W_NAME, W_CODE, W_QTY].map(width => ({ width }));

    const frame = (cell, top, bottom) => {
      cell.border = { top: top || med, bottom: bottom || med, left: thick, right: thick };
    };
    const banner = (row, value, font, height, bg) => {
      ws.mergeCells(row, 1, row, 3);
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
    ['שם המוצר', 'ברקוד', 'כמות'].forEach((h, i) => {
      const c = head.getCell(i + 1);
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
      const code = it.barcodeNum || it.ean || '';
      const row = ws.getRow(r);
      const nameRich = [{ text: (unit ? '★ ' : '') + it.name, font: { name: 'Arial', bold: true, size: NAME_PT } }];
      if (it.en) nameRich.push({ text: '\n' + it.en, font: { name: 'Arial', size: EN_PT, color: { argb: GREY } } });
      row.height = nameLines((unit ? '★ ' : '') + it.name) * NAME_PT * 1.25 + (it.en ? EN_PT * 1.4 : 0) + 12;

      const a = row.getCell(1);
      a.value = { richText: nameRich };
      a.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true, readingOrder: 'rtl' };
      const b = row.getCell(2);
      b.value = code || '—';
      b.font = { name: 'Arial', bold: true, size: CODE_PT, color: { argb: code ? 'FF000000' : GREY } };
      b.alignment = { horizontal: 'center', vertical: 'middle' };
      const c = row.getCell(3);
      c.value = { richText: [{ text: unit ? 'יחידות' : 'ארגזים', font: { name: 'Arial', size: 10, color: { argb: 'FFBDBDBD' } } }] };
      c.alignment = { horizontal: 'center', vertical: 'bottom' };
      if (unit) { a.fill = fill(UNIT_BG); b.fill = fill(UNIT_BG); }
      [a, b, c].forEach(x => frame(x));
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
      [1, 2, 3].forEach(k => frame(row.getCell(k)));
    }
    return wb;
  }

  root.buildAviWorkbook = buildAviWorkbook;
  if (typeof module !== 'undefined' && module.exports) module.exports = { buildAviWorkbook };
})(typeof window !== 'undefined' ? window : globalThis);
