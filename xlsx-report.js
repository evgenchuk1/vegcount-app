/* ── Excel report builder (shared: browser + node) ─────────────────────
   buildCountWorkbook(ExcelJS, sections, meta) → ExcelJS.Workbook
   sections: [{ title, icon, color, rows:[{name, parts:[n], boxes, coeff, kg, extra}] }]
   meta:     { date:'dd.mm.yyyy', time:'HH:MM', branch:'118' }
*/
(function (root) {
  const GREEN = 'FF1B5E20', LIGHT = 'FFE8F5E9', GREY = 'FF9E9E9E', BORDER = 'FFBDBDBD';
  const thin = { style: 'thin', color: { argb: BORDER } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const r1 = n => Math.round(n * 10) / 10;

  function buildCountWorkbook(ExcelJS, sections, meta) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ספירת ירקות ופירות';
    wb.created = new Date();
    const ws = wb.addWorksheet('ספירה', {
      views: [{ rightToLeft: true, state: 'frozen', ySplit: 5 }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    ws.columns = [
      { key: 'n', width: 5 }, { key: 'name', width: 30 }, { key: 'parts', width: 16 },
      { key: 'boxes', width: 11 }, { key: 'coeff', width: 11 }, { key: 'kg', width: 12 },
    ];

    let counted = 0, totalKg = 0;
    sections.forEach(s => s.rows.forEach(r => { if (r.boxes > 0 || (r.extra && r.kg > 0)) { counted++; totalKg += r.kg || 0; } }));

    ws.mergeCells('A1:F1');
    Object.assign(ws.getCell('A1'), { value: `ספירת ירקות ופירות — סניף ${meta.branch}` });
    ws.getCell('A1').font = { bold: true, size: 18, color: { argb: 'FFFFFFFF' } };
    ws.getCell('A1').fill = fill('FF2E7D32');
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 32;

    ws.mergeCells('A2:F2');
    ws.getCell('A2').value = `תאריך: ${meta.date}    שעה: ${meta.time}`;
    ws.getCell('A2').font = { size: 12, color: { argb: GREEN } };
    ws.getCell('A2').alignment = { horizontal: 'center' };

    ws.mergeCells('A3:F3');
    ws.getCell('A3').value = `פריטים שנספרו: ${counted}    |    סה"כ משקל: ${r1(totalKg)} ק"ג`;
    ws.getCell('A3').font = { bold: true, size: 12 };
    ws.getCell('A3').fill = fill(LIGHT);
    ws.getCell('A3').alignment = { horizontal: 'center' };

    const head = ws.getRow(5);
    head.values = ['#', 'מוצר', 'ארגזים (לפי מיקום)', 'סה"כ ארגזים', 'מקדם ק"ג', 'סה"כ ק"ג'];
    head.height = 22;
    head.eachCell(c => {
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = fill('FF388E3C');
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = box;
    });

    let rowNo = 6;
    sections.forEach(sec => {
      if (!sec.rows.length) return;
      ws.mergeCells(rowNo, 1, rowNo, 6);
      const sc = ws.getCell(rowNo, 1);
      sc.value = `${sec.icon} ${sec.title}`;
      sc.font = { bold: true, size: 13, color: { argb: GREEN } };
      sc.fill = fill(sec.color);
      sc.alignment = { horizontal: 'center' };
      ws.getRow(rowNo).height = 20;
      rowNo++;

      const first = rowNo;
      sec.rows.forEach((r, i) => {
        const row = ws.getRow(rowNo);
        const has = r.boxes > 0 || (r.extra && r.kg > 0);
        row.getCell(1).value = i + 1;
        row.getCell(2).value = r.name;
        if (r.extra) {
          row.getCell(6).value = r.kg || null;
        } else {
          row.getCell(3).value = r.parts.length > 1 ? r.parts.join(' + ') : null;
          row.getCell(4).value = r.boxes > 0 ? r.boxes : null;
          row.getCell(5).value = r.coeff || null;
          if (r.boxes > 0 && r.coeff) row.getCell(6).value = { formula: `D${rowNo}*E${rowNo}`, result: r1(r.kg) };
          else if (r.boxes > 0 && /\(יח'?\)/.test(r.name)) row.getCell(6).value = `${r.boxes} יח'`;  // per-unit item
        }
        for (let c = 1; c <= 6; c++) {
          const cell = row.getCell(c);
          cell.border = box;
          cell.alignment = { horizontal: c === 2 ? 'right' : 'center', vertical: 'middle' };
          cell.font = { name: 'Arial', size: 11, color: { argb: has ? 'FF000000' : GREY } };
          if (has) cell.fill = fill(LIGHT);
        }
        row.getCell(6).numFmt = '0.0';
        if (has) {                                   // highlighted total cell
          row.getCell(6).font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF000000' } };
          row.getCell(6).fill = fill('FF00B050');
          row.height = 19;
        }
        rowNo++;
      });
      const sub = ws.getRow(rowNo);
      sub.getCell(5).value = 'סה"כ:';
      sub.getCell(6).value = { formula: `SUM(F${first}:F${rowNo - 1})`, result: r1(sec.rows.reduce((a, r) => a + (r.kg || 0), 0)) };
      sub.getCell(6).numFmt = '0.0';
      [5, 6].forEach(c => {
        sub.getCell(c).font = { bold: true, color: { argb: GREEN } };
        sub.getCell(c).alignment = { horizontal: 'center' };
        sub.getCell(c).border = { top: { style: 'medium', color: { argb: GREEN } } };
      });
      rowNo += 2;
    });
    return wb;
  }

  root.buildCountWorkbook = buildCountWorkbook;
  if (typeof module !== 'undefined') module.exports = { buildCountWorkbook };
})(typeof window !== 'undefined' ? window : globalThis);
