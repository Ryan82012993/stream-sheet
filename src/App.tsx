import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import './index.css';
import LuckyExcel from 'luckyexcel';
import { saveAs } from 'file-saver';
import { exportExcelFile } from './utils/excelExporter';
import { Upload, Download, RefreshCw, AlertCircle, CheckCircle, FolderOpen, FileSpreadsheet, Plus, X } from 'lucide-react';

const URL = 'http://localhost:3001';

interface WorkbookItem {
  id: string;
  fileName: string;
  sheets: any[];
  fileHandle: any | null;
  backendSync: boolean;
  saveStatus: 'idle' | 'saving' | 'saved' | 'error';
  workbookKey: string;
}

const ensureCellData = (sheets: any[]): any[] => {
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
      calcChain = calcChain.filter((item: any) => {
        if (!item || typeof item.r !== 'number' || typeof item.c !== 'number') return false;
        
        // 强制对齐 id 为当前工作表的 ID
        item.id = sheetId;

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

    const updatedSheet: any = {
      ...s,
      id: sheetId,
      index: sheetId,
      celldata,
      calcChain
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

export default function App() {
  const [workbooks, setWorkbooks] = useState<WorkbookItem[]>([
    {
      id: 'sandbox-default',
      fileName: '未命名表格.xlsx',
      sheets: [{ name: 'Sheet1', id: 'sheet-1', index: 'sheet-1', status: 1, celldata: [], order: 0 }],
      fileHandle: null,
      backendSync: false,
      saveStatus: 'idle',
      workbookKey: 'default-key'
    }
  ]);
  const [activeId, setActiveId] = useState<string>('sandbox-default');
  const [err, setErr] = useState<string | null>(null);
  const [backendActive, setBackendActive] = useState(false);
  const [capturedErrors, setCapturedErrors] = useState<{ message: string; stack?: string; time: string }[]>([]);

  const timer = useRef<any>(null);
  const isImportingRef = useRef<boolean>(true);
  const lastLoadTimeRef = useRef<number>(0);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 用 Ref 数组 and Ref ID 保证闭包在定时器等异步回调内拿到最新值，避免经典闭包时序 Bug
  const workbooksRef = useRef<WorkbookItem[]>(workbooks);
  workbooksRef.current = workbooks;
  const activeIdRef = useRef<string>(activeId);
  activeIdRef.current = activeId;

  useEffect(() => {
    const reportError = (message: string, stack?: string) => {
      const time = new Date().toLocaleTimeString();
      setCapturedErrors(prev => [...prev, { message, stack, time }]);
      
      // 异步上联本地日志系统
      fetch(`${URL}/api/log-error`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message,
          stack,
          activeTab: activeIdRef.current
        })
      }).catch(err => console.warn('Failed to send error log to server', err));
    };

    const handleWindowError = (event: ErrorEvent) => {
      const message = event.message || event.error?.message || 'Unknown window error';
      const stack = event.error?.stack;
      reportError(message, stack);
    };

    const handlePromiseRejection = (event: PromiseRejectionEvent) => {
      const message = event.reason?.message || String(event.reason);
      const stack = event.reason?.stack;
      reportError(`[Promise Rejection] ${message}`, stack);
    };

    window.addEventListener('error', handleWindowError);
    window.addEventListener('unhandledrejection', handlePromiseRejection);

    return () => {
      window.removeEventListener('error', handleWindowError);
      window.removeEventListener('unhandledrejection', handlePromiseRejection);
    };
  }, []);

  // 缓存每个 "wbId-wbKey" 组合对应的 stable 回调函数，避免 render 时重建导致 FortuneSheet 无限循环更新
  const changeHandlersRef = useRef<Record<string, (newSheets: any[]) => void>>({});

  const activeWorkbook = workbooks.find(w => w.id === activeId);

  const ensuredSheets = useMemo(() => {
    if (!activeWorkbook) return [];
    return ensureCellData(activeWorkbook.sheets);
  }, [activeWorkbook?.sheets, activeWorkbook?.workbookKey]);

  // 核心：在活动工作簿或其挂载 Key 变更后，延迟解锁。
  // 当切换标签、导入文件重置 workbookKey 挂载渲染并提交到屏幕后，
  // 旧工作表的销毁残留事件（往往伴随多余的 onChange 触发）已彻底在浏览器事件循环中处理完毕，
  // 此时在 useEffect 里安全解锁 isImportingRef。
  useEffect(() => {
    isImportingRef.current = false;
  }, [activeId, activeWorkbook?.workbookKey]);

  const importExcel = (file: File, existingHandle?: any, isBackend: boolean = false) => {
    isImportingRef.current = true;
    if (timer.current) {
      clearTimeout(timer.current);
      timer.current = null;
    }

    try {
      LuckyExcel.transformExcelToLucky(
        file,
        (json: any) => {
          try {
            if (!json.sheets?.length) {
              setErr('格式错误：Excel 文件中未包含有效的工作表');
              return;
            }

            // 规范化工作表状态：有且仅有一个活动工作表 (status = 1)，其余设为 0。防止渲染引擎并发激活导致崩溃。
            let hasActive = false;
            const normalizedSheets = json.sheets.map((s: any) => {
              const isActive = s.status === 1 || s.status === '1';
              const sheetId = s.id || s.index?.toString() || Math.random().toString();
              const finalStatus = isActive ? (hasActive ? 0 : 1) : 0;
              if (finalStatus === 1) {
                hasActive = true;
              }
              return {
                ...s,
                id: sheetId,
                index: sheetId, // 强一致性绑定：确保 index 与 id 恒等，规避 FortuneSheet 计算公式和定位单元格时因 index 缺失或不一致导致的 Cannot read properties of undefined 崩溃
                status: finalStatus
              };
            });

            if (!hasActive && normalizedSheets.length > 0) {
              normalizedSheets[0].status = 1;
            }

            const finalizedSheets = ensureCellData(normalizedSheets);

            // 根据是否是后端同步生成 ID
            const newId = isBackend 
              ? 'backend-sync' 
              : (existingHandle ? 'file-' + Date.now() : 'upload-' + Date.now());

            // 解决同名冲突：在标签页重名时，自动重命名（例如：data (1).xlsx）
            let finalFileName = file.name;
            const existingNames = workbooksRef.current
              .filter(w => {
                if (w.id === 'sandbox-default') {
                  const isDefaultEmpty = w.sheets.length === 1 && 
                    w.sheets[0].name === 'Sheet1' && 
                    (!w.sheets[0].celldata || w.sheets[0].celldata.length === 0);
                  return !isDefaultEmpty;
                }
                return true;
              })
              .map(w => w.fileName);

            if (existingNames.includes(finalFileName)) {
              const extIndex = file.name.lastIndexOf('.');
              const name = extIndex !== -1 ? file.name.substring(0, extIndex) : file.name;
              const ext = extIndex !== -1 ? file.name.substring(extIndex) : '';
              let counter = 1;
              while (existingNames.includes(`${name} (${counter})${ext}`)) {
                counter++;
              }
              finalFileName = `${name} (${counter})${ext}`;
            }

            const newWb: WorkbookItem = {
              id: newId,
              fileName: finalFileName,
              sheets: finalizedSheets,
              fileHandle: existingHandle || null,
              backendSync: isBackend,
              saveStatus: 'saved',
              workbookKey: Math.random().toString()
            };

            setWorkbooks(prev => {
              // 离场数据净化：深拷贝即将变更为不活跃的当前标签页数据，彻底剥离 FortuneSheet 挂载在 sheets 上的 DOM 等脏引用，防止组件重新挂载时崩溃
              const purifiedPrev = prev.map(w => w.id === activeIdRef.current ? { ...w, sheets: JSON.parse(JSON.stringify(w.sheets)) } : w);

              // 如果只剩下一个未命名默认空白表，且内容为空，直接清理掉，避免用户手动关闭的麻烦
              const filtered = purifiedPrev.filter(w => {
                if (w.id === 'sandbox-default') {
                  const isDefaultEmpty = w.sheets.length === 1 && 
                    w.sheets[0].name === 'Sheet1' && 
                    (!w.sheets[0].celldata || w.sheets[0].celldata.length === 0);
                  return !isDefaultEmpty;
                }
                return true;
              });

              // 去重并追加
              const final = filtered.filter(w => w.id !== newId);
              return [...final, newWb];
            });

            setActiveId(newId);
            lastLoadTimeRef.current = Date.now();
            setErr(null);
          } catch (e: any) {
            console.error('Processing sheets failed:', e);
            setErr('表格数据结构规整失败: ' + e.message);
            isImportingRef.current = false;
          }
        },
        (err: any) => {
          console.error('LuckyExcel translation failed:', err);
          const detail = typeof err === 'string' ? err : err?.message || '文件格式可能不兼容或已损坏';
          setErr('解析 Excel 失败: ' + detail);
          isImportingRef.current = false;
        }
      );
    } catch (e: any) {
      console.error('importExcel error:', e);
      setErr('读取 Excel 异常: ' + e.message);
      isImportingRef.current = false;
    }
  };

  // 挂载时检测伴侣流服务器状态
  useEffect(() => {
    fetch(`${URL}/api/status`).then(r => r.json()).then(d => {
      setBackendActive(true);
      fetch(`${URL}/api/load`).then(r => r.blob()).then(b => {
        importExcel(new File([b], d.fileName), null, true);
      }).catch(() => {
        setErr('加载伴侣服务器文件失败');
        isImportingRef.current = false;
      });
    }).catch(() => {
      setBackendActive(false);
      isImportingRef.current = false;
    });
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) {
      importExcel(e.target.files[0]);
      e.target.value = ''; // 允许重复选择同一个同名文件并触发 onChange
    }
  };

  const handleOpenFileHandle = async () => {
    // 兼容性优雅降级：Safari / Firefox 不支持 showOpenFilePicker 时自动回退触发 input 选择器
    if (typeof (window as any).showOpenFilePicker !== 'function') {
      console.warn('showOpenFilePicker is not supported in this browser. Falling back to native file input.');
      fileInputRef.current?.click();
      return;
    }

    try {
      const [handle] = await (window as any).showOpenFilePicker({
        types: [{
          description: 'Excel 工作簿 (*.xlsx, *.xls)',
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel': ['.xlsx', '.xls'],
            'application/msexcel': ['.xlsx', '.xls'],
            'application/x-msexcel': ['.xlsx', '.xls'],
            'application/x-ms-excel': ['.xlsx', '.xls'],
            'application/x-excel': ['.xlsx', '.xls'],
            'application/xls': ['.xlsx', '.xls'],
            'application/x-xls': ['.xlsx', '.xls'],
            'application/octet-stream': ['.xlsx', '.xls']
          }
        }]
      });
      if (handle) {
        // 重复打开同一文件去重检测
        let alreadyOpenWb: WorkbookItem | null = null;
        for (const wb of workbooksRef.current) {
          if (wb.fileHandle) {
            try {
              const isSame = await handle.isSameEntry(wb.fileHandle);
              if (isSame) {
                alreadyOpenWb = wb;
                break;
              }
            } catch (errSame) {
              console.error('比对文件句柄失败:', errSame);
            }
          }
        }

        if (alreadyOpenWb) {
          // 若文件已在标签页中打开，直接聚焦已有标签页，不重复加载
          handleSwitchTab(alreadyOpenWb.id);
          return;
        }

        importExcel(await handle.getFile(), handle);
      }
    } catch (e: any) {
      if (e.name !== 'AbortError') {
        console.error('showOpenFilePicker error:', e);
        setErr('打开文件失败: ' + e.message);
      }
    }
  };

  const performSave = async (targetId: string, currSheets: any[]) => {
    const target = workbooksRef.current.find(w => w.id === targetId);
    if (!target) return;
    if (!target.fileHandle && !target.backendSync) return;

    setWorkbooks(prev => prev.map(w => w.id === targetId ? { ...w, saveStatus: 'saving' } : w));
    try {
      const blob = await exportExcelFile(currSheets);
      if (target.fileHandle) {
        const wr = await target.fileHandle.createWritable();
        await wr.write(blob);
        await wr.close();
      }
      if (target.backendSync && backendActive) {
        await fetch(`${URL}/api/save`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/octet-stream' },
          body: blob
        });
      }
      setWorkbooks(prev => prev.map(w => w.id === targetId ? { ...w, saveStatus: 'saved' } : w));
    } catch (e: any) {
      console.error('Save failed:', e);
      setWorkbooks(prev => prev.map(w => w.id === targetId ? { ...w, saveStatus: 'error' } : w));
      setErr(`保存失败 (${target.fileName}): ` + e.message);
    }
  };

  const flushPendingSave = () => {
    if (timer.current) {
      clearTimeout(timer.current);
      timer.current = null;
      const activeWb = workbooksRef.current.find(w => w.id === activeIdRef.current);
      if (activeWb && (activeWb.fileHandle || (activeWb.backendSync && backendActive))) {
        performSave(activeWb.id, activeWb.sheets);
      }
    }
  };

  const handleDataChangeForWb = useCallback((wbId: string, wbKey: string, newSheets: any[]) => {
    // 0. 前置安全锁：拦截挂载/导入初始化时，由 FortuneSheet 派发的未就绪或空 sheets 数据
    // 防止其将 React 状态中已有的真实/完整数据覆写为空，确保数据不会丢失
    if (isImportingRef.current) {
      console.warn(`[Guard] 成功拦截挂载初始化未就绪期 (${wbId}) 触发的 onChange 事件`);
      return;
    }

    // 1. 安全防护：如果是已经被切走的标签页触发了旧实例的卸载 onChange，直接丢弃
    if (wbId !== activeIdRef.current) {
      console.warn(`[Guard] 成功拦截非活动标签页 (${wbId}) 触发的残留 onChange 事件`);
      return;
    }

    const target = workbooksRef.current.find(w => w.id === wbId);
    // 2. 安全防护：如果当前的 Key 已经发生改变（说明旧实例正在被销毁重构），直接拦截并丢弃
    if (!target || target.workbookKey !== wbKey) {
      console.warn(`[Guard] 成功拦截旧实例 Key (${wbKey}) 触发的销毁期 onChange 事件`);
      return;
    }

    // 校验通过：安全更新对应工作簿的内存状态
    // 为了防止不必要的重新渲染，在状态更新前进行数据引用对比，若无变化则直接阻断以防循环
    setWorkbooks(prev => {
      const affectedWb = prev.find(w => w.id === wbId);
      if (affectedWb && affectedWb.sheets === newSheets) {
        return prev; // 引用完全一致，阻断更新循环
      }
      return prev.map(w => w.id === wbId ? { ...w, sheets: newSheets } : w);
    });

    if (!target.fileHandle && !target.backendSync) return;

    if (Date.now() - lastLoadTimeRef.current < 2500) {
      return;
    }

    setWorkbooks(prev => prev.map(w => w.id === wbId ? { ...w, saveStatus: 'saving' } : w));
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => {
      timer.current = null;
      performSave(wbId, newSheets);
    }, 1500);
  }, [backendActive]);

  // 使用 Ref 引用最新的 handleDataChangeForWb，解决回调函数的 stale closure 问题
  const handleDataChangeForWbRef = useRef(handleDataChangeForWb);
  handleDataChangeForWbRef.current = handleDataChangeForWb;

  // 统一的、引用极其稳定的稳定化 handler 缓存获取函数
  const getChangeHandler = (wbId: string, wbKey: string) => {
    const cacheKey = `${wbId}-${wbKey}`;
    if (!changeHandlersRef.current[cacheKey]) {
      changeHandlersRef.current[cacheKey] = (newSheets: any[]) => {
        handleDataChangeForWbRef.current(wbId, wbKey, newSheets);
      };
    }
    return changeHandlersRef.current[cacheKey];
  };

  useEffect(() => {
    const handleBeforeUnload = () => {
      flushPendingSave();
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, []);

  const handleDownload = async () => {
    const target = workbooksRef.current.find(w => w.id === activeIdRef.current);
    if (!target) return;
    try {
      setWorkbooks(prev => prev.map(w => w.id === target.id ? { ...w, saveStatus: 'saving' } : w));
      saveAs(await exportExcelFile(target.sheets), target.fileName);
      setWorkbooks(prev => prev.map(w => w.id === target.id ? { ...w, saveStatus: 'saved' } : w));
    } catch (e: any) {
      setErr('导出失败');
      setWorkbooks(prev => prev.map(w => w.id === target.id ? { ...w, saveStatus: 'error' } : w));
    }
  };

  const handleSwitchTab = (id: string) => {
    if (id === activeId) return;
    flushPendingSave();

    // 离场数据净化：深拷贝即将变更为不活跃的当前标签页数据，彻底剥离 FortuneSheet 挂载在 sheets 上的 DOM 等脏引用，防止组件重新挂载时崩溃
    if (activeId) {
      setWorkbooks(prev => prev.map(w => w.id === activeId ? { ...w, sheets: JSON.parse(JSON.stringify(w.sheets)) } : w));
    }

    setActiveId(id);
    isImportingRef.current = true;
    lastLoadTimeRef.current = Date.now();
  };

  const handleCloseTab = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    
    if (id === activeId) {
      flushPendingSave();
    }

    setWorkbooks(prev => {
      const filtered = prev.filter(w => w.id !== id);
      if (id === activeId) {
        if (filtered.length > 0) {
          const currentIndex = prev.findIndex(w => w.id === id);
          const nextActiveIndex = Math.max(0, currentIndex - 1);
          const nextWb = filtered[nextActiveIndex];
          setTimeout(() => {
            setActiveId(nextWb.id);
            isImportingRef.current = true;
            lastLoadTimeRef.current = Date.now();
          }, 0);
        } else {
          setTimeout(() => {
            setActiveId('');
          }, 0);
        }
      }
      return filtered;
    });
  };

  const handleNewTab = () => {
    flushPendingSave();

    // 离场数据净化：深拷贝即将变更为不活跃的当前标签页数据，彻底剥离 FortuneSheet 挂载在 sheets 上的 DOM 等脏引用，防止组件重新挂载时崩溃
    if (activeId) {
      setWorkbooks(prev => prev.map(w => w.id === activeId ? { ...w, sheets: JSON.parse(JSON.stringify(w.sheets)) } : w));
    }

    const newId = 'sandbox-new-' + Date.now();
    const newWb: WorkbookItem = {
      id: newId,
      fileName: `新工作簿-${workbooks.length + 1}.xlsx`,
      sheets: [{ name: 'Sheet1', id: 'sheet-1', index: 'sheet-1', status: 1, celldata: [], order: 0 }],
      fileHandle: null,
      backendSync: false,
      saveStatus: 'idle',
      workbookKey: Math.random().toString()
    };
    setWorkbooks(prev => [...prev, newWb]);
    setActiveId(newId);
    isImportingRef.current = true;
    lastLoadTimeRef.current = Date.now();
  };

  const fileName = activeWorkbook ? activeWorkbook.fileName : '无打开的工作簿';
  const saveStatus = activeWorkbook ? activeWorkbook.saveStatus : 'idle';

  return (
    <div className="app-container">
      <header className="app-header">
        <div className="logo-area">
          <div className="excel-logo"><FileSpreadsheet size={20} color="#ffffff" /></div>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '2px' }}>
              <span style={{ fontSize: '10px', background: 'linear-gradient(135deg, #10b981, #059669)', color: '#ffffff', padding: '2px 8px', borderRadius: '4px', fontWeight: 'bold', letterSpacing: '0.5px', textTransform: 'uppercase' }}>StreamSheet</span>
              <h1 className="file-title" style={{ margin: 0 }}>{fileName}</h1>
            </div>
            <p className="file-subtitle">
              {activeWorkbook ? (
                activeWorkbook.fileHandle 
                  ? '💻 File System Access (物理直写)' 
                  : activeWorkbook.backendSync 
                    ? '⚡ Companion Stream Server (伴侣流同步)' 
                    : '☁️ Client Sandbox (离线临时编辑)'
              ) : (
                '请新建或打开 Excel 文件开始编辑'
              )}
            </p>
          </div>
        </div>
        <div className="actions-area">
          <div className="status-label">
            {saveStatus === 'saving' && <span className="status-saving"><RefreshCw size={14} className="animate-spin" /> 保存中</span>}
            {saveStatus === 'saved' && <span className="status-saved"><CheckCircle size={14} /> 已保存</span>}
            {saveStatus === 'error' && <span className="status-error"><AlertCircle size={14} /> 失败</span>}
          </div>
          <button onClick={handleOpenFileHandle} className="btn btn-primary"><FolderOpen size={15} /> 打开Excel</button>
          <label className="btn btn-secondary">
            <Upload size={15} /> 上传
            <input ref={fileInputRef} type="file" accept=".xlsx" onChange={handleFileUpload} style={{ display: 'none' }} />
          </label>
          <button onClick={handleDownload} disabled={!activeWorkbook} className="btn btn-success"><Download size={15} /> 导出</button>
        </div>
      </header>

      {/* 多文件 TabBar */}
      <div className="tabs-bar">
        {workbooks.map(wb => (
          <div
            key={wb.id}
            className={`tab-item ${wb.id === activeId ? 'active' : ''}`}
            onClick={() => handleSwitchTab(wb.id)}
          >
            <span className="tab-icon">
              <FileSpreadsheet size={14} color={wb.id === activeId ? '#107c41' : '#6b7280'} />
            </span>
            <span className="tab-name" title={wb.fileName}>{wb.fileName}</span>
            <button
              className="tab-close"
              onClick={(e) => handleCloseTab(wb.id, e)}
              title="关闭"
            >
              <X size={12} />
            </button>
          </div>
        ))}
        <button className="tab-add-btn" onClick={handleNewTab} title="新建空白工作簿">
          <Plus size={16} />
        </button>
      </div>

      {err && (
        <div className="error-banner">
          <AlertCircle size={16} /> <span>{err}</span>
          <button onClick={() => setErr(null)} className="close-btn">×</button>
        </div>
      )}

      <div className="workbook-container">
        {activeWorkbook ? (
          <Workbook 
            key={activeWorkbook.workbookKey} 
            data={ensuredSheets} 
            onChange={getChangeHandler(activeWorkbook.id, activeWorkbook.workbookKey)} 
          />
        ) : (
          <div className="empty-workbook-state">
            <FileSpreadsheet size={48} color="#9ca3af" />
            <h3 className="empty-title">没有活动的工作簿</h3>
            <p className="empty-subtitle">点击上方“打开Excel” / “上传”，或者点击标签栏的 “+” 按钮新建工作簿</p>
          </div>
        )}
      </div>

      {capturedErrors.length > 0 && (
        <div style={{
          position: 'fixed',
          bottom: 20,
          right: 20,
          width: '420px',
          maxHeight: '300px',
          background: '#1e1e2e',
          color: '#f38ba8',
          border: '1px solid #f38ba8',
          borderRadius: '8px',
          boxShadow: '0 8px 30px rgba(0,0,0,0.5)',
          padding: '16px',
          fontSize: '12px',
          fontFamily: 'monospace',
          zIndex: 99999,
          overflowY: 'auto',
          display: 'flex',
          flexDirection: 'column',
          gap: '8px'
        }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f38ba833', paddingBottom: '6px' }}>
            <span style={{ fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <AlertCircle size={14} /> 捕获到浏览器运行时错误 ({capturedErrors.length})
            </span>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button 
                onClick={() => {
                  navigator.clipboard.writeText(JSON.stringify(capturedErrors, null, 2));
                  alert('错误日志已复制到剪贴板！');
                }} 
                style={{ background: '#313244', color: '#cdd6f4', border: 'none', borderRadius: '4px', padding: '2px 8px', cursor: 'pointer' }}
              >
                复制日志
              </button>
              <button 
                onClick={() => setCapturedErrors([])} 
                style={{ background: '#f38ba8', color: '#11111b', border: 'none', borderRadius: '4px', padding: '2px 8px', cursor: 'pointer', fontWeight: 'bold' }}
              >
                清空
              </button>
            </div>
          </div>
          <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '8px' }}>
            {capturedErrors.map((err, index) => (
              <div key={index} style={{ borderBottom: index < capturedErrors.length - 1 ? '1px dashed #313244' : 'none', paddingBottom: '6px' }}>
                <div style={{ color: '#f9e2af', fontWeight: 'bold', marginBottom: '4px' }}>[{err.time}] {err.message}</div>
                {err.stack && (
                  <pre style={{ margin: 0, padding: '4px', background: '#11111b', borderRadius: '4px', whiteSpace: 'pre-wrap', wordBreak: 'break-all', fontSize: '10px', color: '#a6adc8', maxHeight: '100px', overflowY: 'auto' }}>
                    {err.stack}
                  </pre>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
