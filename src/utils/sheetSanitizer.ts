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
  status?: number | string;
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
      const clonedChain = calcChain.map((item: CalcChainItem) => {
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

    return updatedSheet;
  });
};
