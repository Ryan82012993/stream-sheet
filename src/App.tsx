import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import './index.css';
import LuckyExcel from 'luckyexcel';
import { saveAs } from 'file-saver';
import { exportExcelFile } from './utils/excelExporter';
import { Upload, Download, RefreshCw, AlertCircle, CheckCircle, FolderOpen, FileSpreadsheet, Plus, X, HelpCircle } from 'lucide-react';
import { ensureCellData, SheetData, hasCellDataChanged } from './utils/sheetSanitizer';

const URL = import.meta.env.VITE_COMPANION_SERVER_URL || 'http://localhost:3001';

interface WorkbookItem {
  id: string;
  fileName: string;
  sheets: SheetData[];
  fileHandle: any | null;
  backendSync: boolean;
  saveStatus: 'idle' | 'saving' | 'saved' | 'error';
  workbookKey: string;
}

export default function App() {
  const [workbooks, setWorkbooks] = useState<WorkbookItem[]>([]);
  const [activeId, setActiveId] = useState<string>('');
  const [err, setErr] = useState<string | null>(null);
  const [backendActive, setBackendActive] = useState(false);
  const [capturedErrors, setCapturedErrors] = useState<{ message: string; stack?: string; time: string }[]>([]);
  const [isDragging, setIsDragging] = useState<boolean>(false);

  const timer = useRef<any>(null);
  const isImportingRef = useRef<boolean>(false);
  const lastLoadTimeRef = useRef<number>(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dragCounter = useRef<number>(0);

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
              
              // 关键修复：从 LuckyExcel 导入时，必须主动剔除 s.data！
              // 因为 LuckyExcel 输出的 s.data 往往是缺少合并单元格（mc）等元数据的骨架二维数组。
              // 如果保留 s.data，ensureCellData 会基于此空骨架覆写 celldata，且 FortuneSheet 也会优先使用它而忽略富文本的 s.celldata。
              // 设为 undefined 后，FortuneSheet 会自动从富文本的 s.celldata + s.config.merge 完美生成 2D 矩阵并高保真渲染合并单元格。
              const { data, ...rest } = s;

              return {
                ...rest,
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
              const purifiedPrev = prev.map(w => w.id === activeIdRef.current ? { ...w, sheets: structuredClone(w.sheets) } : w);

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
    fetch(`${URL}/api/status`).then(r => r.json()).then(() => {
      setBackendActive(true);
    }).catch(() => {
      setBackendActive(false);
    });
  }, []);

  // 组件卸载时，物理清除防抖自动保存计时器，防止热重载或跳转时内存泄露引发状态更新报错
  useEffect(() => {
    return () => {
      if (timer.current) {
        clearTimeout(timer.current);
        timer.current = null;
      }
    };
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) {
      importExcel(e.target.files[0]);
      e.target.value = ''; // 允许重复选择同一个同名文件并触发 onChange
    }
  };

  const handleOpenFileHandle = async () => {
    // 兼容性优雅降级：Safari / Firefox 不支持 showOpenFilePicker 时自动回退触发 input 选择器
    if (typeof window.showOpenFilePicker !== 'function') {
      console.warn('showOpenFilePicker is not supported in this browser. Falling back to native file input.');
      fileInputRef.current?.click();
      return;
    }

    try {
      const [handle] = await window.showOpenFilePicker({
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

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.types.includes('Files')) {
      dragCounter.current++;
      setIsDragging(true);
    }
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    dragCounter.current--;
    if (dragCounter.current <= 0) {
      dragCounter.current = 0;
      setIsDragging(false);
    }
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    dragCounter.current = 0;
    setIsDragging(false);

    const items = [...e.dataTransfer.items];
    const files = [...e.dataTransfer.files];

    if (files.length === 0) return;

    const excelFile = files.find(f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));
    if (!excelFile) {
      setErr('不支持的文件格式：请拖入 .xlsx 或 .xls 格式的 Excel 工作簿');
      return;
    }

    let fileHandle: any = null;
    const excelItem = items.find(item => item.kind === 'file' && (item.type.includes('sheet') || item.type.includes('excel') || excelFile.name.endsWith('.xlsx') || excelFile.name.endsWith('.xls')));
    
    if (excelItem && typeof (excelItem as any).getAsFileSystemHandle === 'function') {
      try {
        const handle = await (excelItem as any).getAsFileSystemHandle();
        if (handle && handle.kind === 'file') {
          fileHandle = handle;
        }
      } catch (errHandle) {
        console.warn('获取拖拽文件句柄失败，将降级为只读沙盒模式:', errHandle);
      }
    }

    if (fileHandle) {
      let alreadyOpenWb: WorkbookItem | null = null;
      for (const wb of workbooksRef.current) {
        if (wb.fileHandle) {
          try {
            const isSame = await fileHandle.isSameEntry(wb.fileHandle);
            if (isSame) {
              alreadyOpenWb = wb;
              break;
            }
          } catch (errSame) {
            console.error('拖拽比对文件句柄失败:', errSame);
          }
        }
      }

      if (alreadyOpenWb) {
        handleSwitchTab(alreadyOpenWb.id);
        return;
      }

      try {
        const file = await fileHandle.getFile();
        importExcel(file, fileHandle);
      } catch (e: any) {
        console.error('读取拖拽文件句柄内容失败，降级为沙盒模式:', e);
        importExcel(excelFile);
      }
    } else {
      importExcel(excelFile);
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

    // 性能优化防御：利用 hasCellDataChanged 高效判断当前操作是否包含实质性的数据（单元格值、公式、样式、行高、列宽、合并设置等）修改
    // 如果没有实质变动（仅因单元格点击、选区聚焦、界面滚动等纯视图状态触发 onChange 派发），则静默跳过自动物理落盘，避免昂贵的 CPU 及硬盘 I/O
    const changed = hasCellDataChanged(target.sheets, newSheets);
    if (!changed) {
      return;
    }

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
      setWorkbooks(prev => prev.map(w => w.id === activeId ? { ...w, sheets: structuredClone(w.sheets) } : w));
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

    // 1. 同步计算过滤后的工作簿列表
    const filtered = workbooks.filter(w => w.id !== id);
    setWorkbooks(filtered);

    // 2. 如果关闭的是当前活动工作簿，同步计算并设置新的 activeId，避免 state updater 回调内副作用
    if (id === activeId) {
      if (filtered.length > 0) {
        const currentIndex = workbooks.findIndex(w => w.id === id);
        const nextActiveIndex = Math.max(0, currentIndex - 1);
        const nextWb = filtered[nextActiveIndex];
        setActiveId(nextWb.id);
        isImportingRef.current = true;
        lastLoadTimeRef.current = Date.now();
      } else {
        setActiveId('');
      }
    }
  };

  const handleNewTab = () => {
    flushPendingSave();

    // 离场数据净化：深拷贝即将变更为不活跃的当前标签页数据，彻底剥离 FortuneSheet 挂载在 sheets 上的 DOM 等脏引用，防止组件重新挂载时崩溃
    if (activeId) {
      setWorkbooks(prev => prev.map(w => w.id === activeId ? { ...w, sheets: structuredClone(w.sheets) } : w));
    }

    const newId = 'sandbox-new-' + Date.now();
    const newWb: WorkbookItem = {
      id: newId,
      fileName: `新工作簿-${workbooks.length + 1}.xlsx`,
      sheets: [{ 
        name: 'Sheet1', 
        id: 'sheet-1', 
        index: 'sheet-1', 
        status: 1, 
        order: 0,
        row: 84,
        column: 30,
        celldata: [],
        config: {}
      }],
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
    <div 
      className="app-container"
      onDragEnter={handleDragEnter}
      onDragOver={(e) => e.preventDefault()}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {isDragging && (
        <div className="drag-drop-overlay">
          <div className="drag-drop-box">
            <FileSpreadsheet size={48} className="animate-bounce" color="#107c41" />
            <h3>释放鼠标即可打开 Excel</h3>
            <p>支持物理直写（限 Chrome / Edge）或离线沙盒安全导入</p>
          </div>
        </div>
      )}
      <header className="app-header">
        <div className="logo-area">
          <div className="excel-logo"><FileSpreadsheet size={20} color="#ffffff" /></div>
          <div>
            <div className="header-brand-group">
              <span className="brand-badge">StreamSheet</span>
              <h1 className="file-title file-title-margin-reset">{fileName}</h1>
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
            <input ref={fileInputRef} type="file" accept=".xlsx" onChange={handleFileUpload} className="hidden-file-input" />
          </label>
          <button onClick={handleDownload} disabled={!activeWorkbook} className="btn btn-success"><Download size={15} /> 导出</button>

          <div className="help-tooltip-container" tabIndex={0}>
            <HelpCircle size={18} className="help-icon" />
            <div className="help-tooltip-popup">
              <h4 className="help-tooltip-title">💡 功能选择与保存说明</h4>
              <div className="help-tooltip-grid">
                <div className="help-tooltip-section">
                  <h5>💻 打开 Excel (本地物理直连)</h5>
                  <p className="help-tooltip-desc">
                    <strong>实时自动保存</strong>：基于浏览器文件系统接口，您的修改将在 1.5 秒内<strong>直接写回本机原 Excel 文件</strong>，无需重新手动下载导出。
                  </p>
                </div>
                <div className="help-tooltip-section">
                  <h5 className="upload-title">☁️ 上传 (网页离线沙盒)</h5>
                  <p className="help-tooltip-desc">
                    <strong>单次只读加载</strong>：文件作为一次性副本读入浏览器沙盒，任何编辑均不会改写您的本机原文件。保存修改需手动点击 <strong>“导出”</strong> 重新下载。
                  </p>
                </div>
              </div>
            </div>
          </div>
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
        <div className="error-console-overlay">
          <div className="error-console-header">
            <span className="error-console-title">
              <AlertCircle size={14} /> 捕获到浏览器运行时错误 ({capturedErrors.length})
            </span>
            <div className="error-console-btn-group">
              <button 
                onClick={() => {
                  navigator.clipboard.writeText(JSON.stringify(capturedErrors, null, 2));
                  alert('错误日志已复制到剪贴板！');
                }} 
                className="error-console-btn error-console-btn-copy"
              >
                复制日志
              </button>
              <button 
                onClick={() => setCapturedErrors([])} 
                className="error-console-btn error-console-btn-clear"
              >
                清空
              </button>
            </div>
          </div>
          <div className="error-console-list">
            {capturedErrors.map((err, index) => (
              <div key={index} className="error-console-item">
                <div className="error-console-item-header">[{err.time}] {err.message}</div>
                {err.stack && (
                  <pre className="error-console-stack">
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
