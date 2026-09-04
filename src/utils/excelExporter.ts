import ExcelJS from 'exceljs';

function parseColorToHex(color: string): string | null {
  if (!color) return null;
  const trimmed = color.trim().toLowerCase();
  if (trimmed.startsWith('#')) {
    let hex = trimmed.substring(1);
    if (hex.length === 3) hex = hex.split('').map(c => c + c).join('');
    return 'FF' + hex.toUpperCase();
  }
  if (trimmed.startsWith('rgb')) {
    const match = trimmed.match(/\d+/g);
    if (match && match.length >= 3) {
      const r = parseInt(match[0], 10).toString(16).padStart(2, '0');
      const g = parseInt(match[1], 10).toString(16).padStart(2, '0');
      const b = parseInt(match[2], 10).toString(16).padStart(2, '0');
      return 'FF' + (r + g + b).toUpperCase();
    }
  }
  const cssColors: Record<string, string> = {
    white: 'FFFFFFFF', black: 'FF000000', red: 'FFFF0000', green: 'FF008000',
    blue: 'FF0000FF', yellow: 'FFFFFF00', orange: 'FFFFA500', purple: 'FF800080',
    gray: 'FF808080', grey: 'FF808080', silver: 'FFC0C0C0', gold: 'FFFFD700',
    pink: 'FFFFC0CB', cyan: 'FF00FFFF', magenta: 'FFFF00FF', brown: 'FFA52A2A'
  };
  return cssColors[trimmed] || null;
}

function applyCellData(worksheet: ExcelJS.Worksheet, r: number, c: number, cellVal: any) {
  if (!cellVal) return;
  const excelCell = worksheet.getCell(r + 1, c + 1);

  // 1. 写入单元格的值或公式
  if (cellVal.f) {
    let formulaStr = cellVal.f;
    if (formulaStr.startsWith('=')) formulaStr = formulaStr.substring(1);
    excelCell.value = {
      formula: formulaStr,
      result: cellVal.v !== undefined ? cellVal.v : null
    };
  } else if (cellVal.v !== undefined && cellVal.v !== null) {
    excelCell.value = cellVal.v;
  }

  // 2. 写入样式字体
  const font: Partial<ExcelJS.Font> = {};
  if (cellVal.bl === 1) font.bold = true;
  if (cellVal.it === 1) font.italic = true;
  if (cellVal.cl === 1) font.strike = true;
  if (cellVal.un === 1) font.underline = true;
  if (cellVal.fs) font.size = Number(cellVal.fs);
  if (cellVal.ff) font.name = String(cellVal.ff);
  if (cellVal.fc) {
    const hexColor = parseColorToHex(cellVal.fc);
    if (hexColor) font.color = { argb: hexColor };
  }
  if (Object.keys(font).length > 0) excelCell.font = font;

  // 3. 写入对齐方式
  const alignment: Partial<ExcelJS.Alignment> = {};
  if (cellVal.ht !== undefined) {
    const htMap: Record<number | string, ExcelJS.Alignment['horizontal']> = {
      0: 'center', 1: 'left', 2: 'right'
    };
    alignment.horizontal = htMap[cellVal.ht];
  }
  if (cellVal.vt !== undefined) {
    const vtMap: Record<number | string, ExcelJS.Alignment['vertical']> = {
      0: 'middle', 1: 'top', 2: 'bottom'
    };
    alignment.vertical = vtMap[cellVal.vt];
  }
  if (cellVal.tb === 2 || cellVal.tb === '2') alignment.wrapText = true;
  if (Object.keys(alignment).length > 0) excelCell.alignment = alignment;

  // 4. 写入背景颜色
  if (cellVal.bg) {
    const hexColor = parseColorToHex(cellVal.bg);
    if (hexColor) {
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: hexColor }
      };
    }
  }

  // 5. 写入数字格式化形式
  if (cellVal.ct && cellVal.ct.fa) {
    excelCell.numFmt = cellVal.ct.fa;
  }
}

export async function exportExcelFile(sheets: any[]): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();

  for (const sheet of sheets) {
    // 修复 Bug：移除 if (sheet.status === 0) continue;
    // 确保所有标签页（工作表）在导出或自动保存时全量保留，防止静默数据丢失。
    const worksheet = workbook.addWorksheet(sheet.name);

    // 如果该工作表在 FortuneSheet 中被隐藏，在导出的 ExcelJS 对象中同样设为隐藏状态
    if (sheet.hide === 1 || sheet.hide === '1') {
      worksheet.state = 'hidden';
    }

    worksheet.views = [{ showGridLines: true }];

    // 1. 设置行高 (px * 0.75)
    if (sheet.config?.rowlen) {
      Object.entries(sheet.config.rowlen).forEach(([rStr, height]) => {
        const r = parseInt(rStr, 10);
        worksheet.getRow(r + 1).height = Number(height) * 0.75;
      });
    }

    // 2. 设置列宽 (px / 8)
    if (sheet.config?.columnlen) {
      Object.entries(sheet.config.columnlen).forEach(([cStr, width]) => {
        const c = parseInt(cStr, 10);
        worksheet.getColumn(c + 1).width = Number(width) / 8;
      });
    }

    // 3. 处理合并单元格
    if (sheet.config?.merge) {
      Object.values(sheet.config.merge).forEach((m: any) => {
        const startRow = m.r + 1;
        const startCol = m.c + 1;
        const endRow = m.r + m.rs;
        const endCol = m.c + m.cs;
        try {
          worksheet.mergeCells(startRow, startCol, endRow, endCol);
        } catch (err) {
          console.warn('Merge cells failed:', err);
        }
      });
    }

    // 4. 填充数据 (支持 data 2D数组 或 celldata 扁平数组)
    if (sheet.data && sheet.data.length > 0) {
      for (let r = 0; r < sheet.data.length; r++) {
        const rowData = sheet.data[r];
        if (!rowData) continue;
        for (let c = 0; c < rowData.length; c++) {
          applyCellData(worksheet, r, c, rowData[c]);
        }
      }
    } else if (sheet.celldata) {
      sheet.celldata.forEach((item: any) => {
        applyCellData(worksheet, item.r, item.c, item.v);
      });
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
