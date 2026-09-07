/**
 * StreamSheet 表格数据清洗与安全规整工具层 (State Sanitization Layer)
 * 
 * 核心职责：
 * 1. 自动对齐工作表的 id 与 index。
 * 2. 在 FortuneSheet 卸载/重挂载时，从二维元数据 data 中动态重构一维 celldata。
 *    防止 FortuneSheet 内部数据缩减为默认尺寸 (84x60)，保持表格的真实行列物理边界。
 * 3. 强力规整并过滤公式链 (calcChain)，排除未定义或 dangling 引用，深拷贝 item 规避只读冻结属性修改 Crash。
 * 4. 自动计算行数（row）与列数（column）的安全下界，规避越界和运行时闪退。
 */

export interface CalcChainItem {
  r: number;
  c: number;
  id?: string;
  [key: string]: any;
}

export interface SheetData {
  id?: string;
  index?: string | number;
  name: string;
  status?: number;
  celldata?: any[];
  data?: any[][];
  calcChain?: CalcChainItem[];
  row?: number;
  column?: number;
  [key: string]: any;
}

/**
 * 强力规整并清洗工作表数组，输出高度安全、对齐的全新副本数据
 * @param sheets 待规整的原始工作表数组
 * @returns 经过清洗后的安全工作表数组
 */
export const ensureCellData = (sheets: SheetData[]): SheetData[] => {
  if (!sheets) return [];

  return sheets.map(s => {
    // 确保 id 与 index 统一且一致
    const sheetId = s.id || s.index?.toString() || Math.random().toString();
    
    // 1. 如果有 s.data，说明处于编辑中或保存状态，通过 s.data 重新构建 celldata
    // 保证在 FortuneSheet 重新挂载初始化时，能够计算出正确的 lastRowNum / lastColNum，防止其缩水为默认尺寸 84x60
    let celldata = s.celldata || [];
    let maxR = 0;
    let maxC = 0;

    if (s.data && s.data.length > 0) {
      const reconstructed: any[] = [];
      for (let r = 0; r < s.data.length; r++) {
        const rowData = s.data[r];
        if (!rowData) continue;
        for (let c = 0; c < rowData.length; c++) {
          const cell = rowData[c];
          if (cell !== null && cell !== undefined) {
            reconstructed.push({ r, c, v: cell });
            if (r > maxR) maxR = r;
            if (c > maxC) maxC = c;
          }
        }
      }
      celldata = reconstructed;
    } else if (celldata.length > 0) {
      celldata.forEach((d: any) => {
        if (d.r > maxR) maxR = d.r;
        if (d.c > maxC) maxC = d.c;
      });
    }

    // 2. 强力规整并过滤公式链（calcChain），过滤掉非公式格或者越界的 dangling references，规避 getcellFormula 闪崩溃
    let calcChain = s.calcChain;
    if (calcChain && Array.isArray(calcChain)) {
      // 通过 map 浅拷贝每一个 item，使其摆脱 Read-only/Frozen 锁定状态变得完全可写，并强制对齐 ID
      const clonedChain = calcChain.map((item: CalcChainItem): CalcChainItem | null => {
        if (!item) return null;
        return { ...item, id: sheetId };
      });

      calcChain = clonedChain.filter((item: CalcChainItem | null): item is CalcChainItem => {
        if (!item || typeof item.r !== 'number' || typeof item.c !== 'number') return false;

        // 如果 s.data 存在，校验对应格是否真的有公式定义
        if (s.data && s.data[item.r] && s.data[item.r][item.c]) {
          const cell = s.data[item.r][item.c];
          return cell && typeof cell === 'object' && !!cell.f;
        }

        // 如果只有 celldata 存在，校验对应 celldata 格是否真的有公式定义
        if (celldata && celldata.length > 0) {
          const hasCellWithFormula = celldata.some((d: any) => d.r === item.r && d.c === item.c && d.v && typeof d.v === 'object' && !!d.v.f);
          return hasCellWithFormula;
        }

        return false;
      });
    }

    const updatedSheet: SheetData = {
      ...s,
      id: sheetId,
      index: sheetId,
      celldata,
      calcChain: calcChain || undefined
    };

    const computedRow = Math.max(s.data?.length || 0, maxR + 1);
    const computedCol = Math.max(s.data?.[0]?.length || 0, maxC + 1);

    // 3. 动态拓展并保护安全边界，防止 initSheetData 时行列数缩水导致的数据丢失或越界崩溃
    if (s.row !== undefined || computedRow > 0) {
      updatedSheet.row = Math.max(s.row || 0, computedRow);
    }
    if (s.column !== undefined || computedCol > 0) {
      updatedSheet.column = Math.max(s.column || 0, computedCol);
    }

    return enrichSheetMerge(updatedSheet);
  });
};


/**
 * 辅助比对：判断单个单元格对象的核心数据与格式是否实质相等
 */
function isCellEqual(c1: any, c2: any): boolean {
  if (c1 === c2) return true;
  if (!c1 || !c2) {
    const empty1 = !c1 || (c1.v === undefined && c1.f === undefined);
    const empty2 = !c2 || (c2.v === undefined && c2.f === undefined);
    return empty1 && empty2;
  }
  
  if (c1.v !== c2.v) return false;
  if (c1.f !== c2.f) return false;
  if (c1.m !== c2.m) return false;
  if (c1.bg !== c2.bg) return false;
  if (c1.fc !== c2.fc) return false;
  if (c1.bl !== c2.bl) return false;
  if (c1.it !== c2.it) return false;
  if (c1.un !== c2.un) return false;
  if (c1.cl !== c2.cl) return false;
  if (c1.fs !== c2.fs) return false;
  if (c1.ff !== c2.ff) return false;
  if (c1.ht !== c2.ht) return false;
  if (c1.vt !== c2.vt) return false;
  
  if (c1.ct?.fa !== c2.ct?.fa) return false;
  if (c1.ct?.t !== c2.ct?.t) return false;
  
  if (c1.mc?.r !== c2.mc?.r || c1.mc?.c !== c2.mc?.c || c1.mc?.rs !== c2.mc?.rs || c1.mc?.cs !== c2.mc?.cs) return false;

  return true;
}

/**
 * 核心优化：判断两个工作簿的数据是否在内容或结构上发生了实质性改动
 * 如果没有改动（仅由于光标移动、格子选中等状态触发的 onChange），返回 false
 */
export function hasCellDataChanged(oldSheets: SheetData[], newSheets: SheetData[]): boolean {
  if (!oldSheets || !newSheets) return true;
  if (oldSheets.length !== newSheets.length) return true;

  for (let i = 0; i < oldSheets.length; i++) {
    const o = oldSheets[i];
    const n = newSheets[i];

    // 1. 基础工作表属性变更
    if (o.name !== n.name) return true;
    if (o.id !== n.id) return true;
    if (o.status !== n.status) return true;
    if (o.hide !== n.hide) return true;
    if (o.row !== n.row || o.column !== n.column) return true;

    // 2. 比对二维数据 data (主要编辑区数据)
    const oData = o.data;
    const nData = n.data;
    if (oData && nData) {
      if (oData.length !== nData.length) return true;
      for (let r = 0; r < oData.length; r++) {
        const oRow = oData[r] || [];
        const nRow = nData[r] || [];
        if (oRow.length !== nRow.length) return true;
        for (let c = 0; c < oRow.length; c++) {
          if (!isCellEqual(oRow[c], nRow[c])) return true;
        }
      }
    } else if (oData || nData) {
      // 一个有 2D data 一个没有，属于重大结构改变
      return true;
    }

    // 3. 比对一维数据 celldata
    const oCelldata = o.celldata || [];
    const nCelldata = n.celldata || [];
    if (oCelldata.length !== nCelldata.length) return true;
    
    // 利用 Map 提升检索效率为 O(1)
    const oCellMap = new Map<string, any>();
    oCelldata.forEach(item => {
      if (item) oCellMap.set(`${item.r}_${item.c}`, item.v);
    });
    for (let j = 0; j < nCelldata.length; j++) {
      const item = nCelldata[j];
      if (!item) continue;
      const key = `${item.r}_${item.c}`;
      if (!oCellMap.has(key)) return true;
      if (!isCellEqual(oCellMap.get(key), item.v)) return true;
    }

    // 4. 比对合并单元格配置 config.merge
    const oMerge = o.config?.merge || {};
    const nMerge = n.config?.merge || {};
    const oMergeKeys = Object.keys(oMerge);
    const nMergeKeys = Object.keys(nMerge);
    if (oMergeKeys.length !== nMergeKeys.length) return true;
    for (const k of oMergeKeys) {
      const om = oMerge[k];
      const nm = nMerge[k];
      if (!nm) return true;
      if (om.r !== nm.r || om.c !== nm.c || om.rs !== nm.rs || om.cs !== nm.cs) return true;
    }

    // 5. 比对行高配置 rowlen
    const oRowlen = o.config?.rowlen || {};
    const nRowlen = n.config?.rowlen || {};
    const oRowlenKeys = Object.keys(oRowlen);
    const nRowlenKeys = Object.keys(nRowlen);
    if (oRowlenKeys.length !== nRowlenKeys.length) return true;
    for (const k of oRowlenKeys) {
      if (oRowlen[k] !== nRowlen[k]) return true;
    }

    // 6. 比对列宽配置 columnlen
    const oCollen = o.config?.columnlen || {};
    const nCollen = n.config?.columnlen || {};
    const oCollenKeys = Object.keys(oCollen);
    const nCollenKeys = Object.keys(nCollen);
    if (oCollenKeys.length !== nCollenKeys.length) return true;
    for (const k of oCollenKeys) {
      if (oCollen[k] !== nCollen[k]) return true;
    }
  }

  return false;
}


/**
 * 核心优化补载：高保真填充合并单元格底层的 mc 属性
 * 1. 若存在 2D data，确保合并区内的每个格子都挂载 mc 指向，左上角格子附带完整的 rs, cs。
 * 2. 若存在 1D celldata，在 Map 中进行校验，如果合并范围内的某些格子因为 Excel 空白被忽略，
 *    强行补全并追加入 celldata 列表中，同时注入 mc 配置，保证 FortuneSheet 的 init 渲染绝对完整。
 */
export function enrichSheetMerge(s: SheetData): SheetData {
  const merge = s.config?.merge;
  if (!merge) return s;

  const sheet = { ...s };

  // 1. 如果存在 2D data，进行单元格级别 mc 注入与补全
  if (sheet.data && sheet.data.length > 0) {
    const data = sheet.data.map(row => (row ? [...row] : []));
    Object.values(merge).forEach((m: any) => {
      if (!m || typeof m.r !== 'number' || typeof m.c !== 'number') return;
      const rs = m.rs || 1;
      const cs = m.cs || 1;

      for (let ri = 0; ri < rs; ri++) {
        const r = m.r + ri;
        if (data.length <= r) continue;
        if (!data[r]) data[r] = [];
        
        for (let ci = 0; ci < cs; ci++) {
          const c = m.c + ci;
          let cell = data[r][c];
          if (cell === null || cell === undefined) {
            cell = {};
          } else if (typeof cell !== 'object') {
            cell = { v: cell };
          } else {
            cell = { ...cell };
          }

          if (ri === 0 && ci === 0) {
            cell.mc = { r: m.r, c: m.c, rs, cs };
          } else {
            cell.mc = { r: m.r, c: m.c };
          }
          data[r][c] = cell;
        }
      }
    });
    sheet.data = data;
  }

  // 2. 如果存在 1D celldata，进行补全并强制注入 mc
  if (sheet.celldata) {
    const celldataMap = new Map<string, any>();
    sheet.celldata.forEach(item => {
      if (item) celldataMap.set(`${item.r}_${item.c}`, item);
    });

    Object.values(merge).forEach((m: any) => {
      if (!m || typeof m.r !== 'number' || typeof m.c !== 'number') return;
      const rs = m.rs || 1;
      const cs = m.cs || 1;

      for (let ri = 0; ri < rs; ri++) {
        const r = m.r + ri;
        for (let ci = 0; ci < cs; ci++) {
          const c = m.c + ci;
          const key = `${r}_${c}`;
          let item = celldataMap.get(key);
          
          if (!item) {
            // 强力补全！原 Excel 中因为空白被省略的合并区域格子，强行追加入 celldata 列表
            item = { r, c, v: {} };
            celldataMap.set(key, item);
          } else {
            const val = item.v === null || item.v === undefined 
              ? {} 
              : (typeof item.v === 'object' ? { ...item.v } : { v: item.v });
            item = { ...item, v: val };
            celldataMap.set(key, item);
          }

          if (ri === 0 && ci === 0) {
            item.v.mc = { r: m.r, c: m.c, rs, cs };
          } else {
            item.v.mc = { r: m.r, c: m.c };
          }
        }
      }
    });

    sheet.celldata = Array.from(celldataMap.values());
  }

  return sheet;
}
