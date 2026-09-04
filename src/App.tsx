import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import './index.css';
import LuckyExcel from 'luckyexcel';
import { saveAs } from 'file-saver';
import { exportExcelFile } from './utils/excelExporter';
import { Upload, Download, RefreshCw, AlertCircle, CheckCircle, FolderOpen, FileSpreadsheet } from 'lucide-react';

const URL = 'http://localhost:3001';

export default function App() {
  const [sheets, setSheets] = useState<any[]>([{ name: 'Sheet1', id: 'sheet-1', status: 1, celldata: [], order: 0 }]);
  const [fileName, setFileName] = useState('未命名表格.xlsx');
  const [fileHandle, setFileHandle] = useState<any>(null);
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle');
  const [err, setErr] = useState<string | null>(null);
  const [backendActive, setBackendActive] = useState(false);
  const timer = useRef<any>(null);
  const sheetsRef = useRef<any[]>(sheets);
  sheetsRef.current = sheets;

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [workbookKey, setWorkbookKey] = useState<string>(Math.random().toString());
  const lastLoadTimeRef = useRef<number>(0);
  const isImportingRef = useRef<boolean>(true);

  // 核心：使用 React 声明周期（useEffect）延迟解锁。
  // 当 workbookKey 改变（代表新 Workbook 挂载渲染并提交到屏幕）后，
  // 此时旧工作表的销毁残留事件已彻底处理完毕，在此处安全解锁 isImportingRef。
  useEffect(() => {
    isImportingRef.current = false;
  }, [workbookKey]);

  const importExcel = (file: File) => {
    isImportingRef.current = true;
    if (timer.current) {
      clearTimeout(timer.current);
      timer.current = null;
    }
    setSaveStatus('saving');
    try {
      LuckyExcel.transformExcelToLucky(
        file,
        (json: any) => {
          try {
            if (!json.sheets?.length) {
              setErr('格式错误：Excel 文件中未包含有效的工作表');
              setSaveStatus('error');
              return;
            }

            // 规范化工作表状态：有且仅有一个活动工作表 (status = 1)，其余设为 0。防止渲染引擎并发激活导致崩溃。
            let hasActive = false;
            const normalizedSheets = json.sheets.map((s: any) => {
              const isActive = s.status === 1 || s.status === '1';
              if (isActive) {
                if (hasActive) {
                  return {
                    ...s,
                    id: s.id || s.index?.toString() || Math.random().toString(),
                    status: 0
                  };
                }
                hasActive = true;
              }
              return {
                ...s,
                id: s.id || s.index?.toString() || Math.random().toString(),
                status: isActive ? 1 : 0
              };
            });

            if (!hasActive && normalizedSheets.length > 0) {
              normalizedSheets[0].status = 1;
            }

            setSheets(normalizedSheets);
            setFileName(file.name);
            setWorkbookKey(Math.random().toString());
            lastLoadTimeRef.current = Date.now();
            setSaveStatus('saved');
            setErr(null);
          } catch (e: any) {
            console.error('Processing sheets failed:', e);
            setErr('表格数据结构规整失败: ' + e.message);
            setSaveStatus('error');
            isImportingRef.current = false;
          }
        },
        (err: any) => {
          console.error('LuckyExcel translation failed:', err);
          const detail = typeof err === 'string' ? err : err?.message || '文件格式可能不兼容或已损坏';
          setErr('解析 Excel 失败: ' + detail);
          setSaveStatus('error');
          isImportingRef.current = false;
        }
      );
    } catch (e: any) {
      console.error('importExcel error:', e);
      setErr('读取 Excel 异常: ' + e.message);
      setSaveStatus('error');
      isImportingRef.current = false;
    }
  };

  useEffect(() => {
    fetch(`${URL}/api/status`).then(r => r.json()).then(d => {
      setBackendActive(true);
      setFileName(d.fileName);
      setSaveStatus('saving');
      fetch(`${URL}/api/load`).then(r => r.blob()).then(b => importExcel(new File([b], d.fileName)))
        .catch(() => {
          setErr('加载失败');
          setSaveStatus('error');
          isImportingRef.current = false;
        });
    }).catch(() => {
      setBackendActive(false);
      isImportingRef.current = false;
    });
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) {
      setFileHandle(null);
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
        setFileHandle(handle);
        importExcel(await handle.getFile());
      }
    } catch (e: any) {
      if (e.name !== 'AbortError') {
        console.error('showOpenFilePicker error:', e);
        setErr('打开文件失败: ' + e.message);
      }
    }
  };

  const performSave = async (curr: any[]) => {
    if (!fileHandle && !backendActive) return;
    setSaveStatus('saving');
    try {
      const blob = await exportExcelFile(curr);
      if (fileHandle) {
        const wr = await fileHandle.createWritable();
        await wr.write(blob);
        await wr.close();
      }
      if (backendActive) {
        await fetch(`${URL}/api/save`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/octet-stream' },
          body: blob
        });
      }
      setSaveStatus('saved');
    } catch (e: any) {
      setSaveStatus('error');
      setErr('保存失败: ' + e.message);
    }
  };

  const handleDataChange = useCallback((newSheets: any[]) => {
    // 1. 无论何种时序下，只要渲染层（FortuneSheet）发出数据更新，都无条件更新 React 状态。
    // 这保证了 Excel 在初始化或挂载过程中补全的数据，能完美、实时地同步回 React 层，解决“加载读取不到”的问题。
    setSheets(newSheets);

    if (!fileHandle && !backendActive) return;

    // 2. 隔离保存：只要是在导入期间，或者是在加载文件的 2.5 秒安全期内，我们都仅仅拦截“自动保存定时器”的触发，
    // 决不触发物理磁盘写入。这样既能保证 0 弹窗打扰，又能保证数据 100% 完整读取。
    if (isImportingRef.current) {
      return;
    }
    if (Date.now() - lastLoadTimeRef.current < 2500) {
      return;
    }

    setSaveStatus('saving');
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => performSave(newSheets), 1500);
  }, [fileHandle, backendActive]);

  const handleDownload = async () => {
    try {
      setSaveStatus('saving');
      saveAs(await exportExcelFile(sheetsRef.current), fileName);
      setSaveStatus('saved');
    } catch (e: any) {
      setErr('导出失败');
      setSaveStatus('error');
    }
  };

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
              {fileHandle ? '💻 File System Access (物理直写)' : backendActive ? '⚡ Companion Stream Server (伴侣流同步)' : '☁️ Client Sandbox (离线临时编辑)'}
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
          <button onClick={handleDownload} className="btn btn-success"><Download size={15} /> 导出</button>
        </div>
      </header>
      {err && (
        <div className="error-banner">
          <AlertCircle size={16} /> <span>{err}</span>
          <button onClick={() => setErr(null)} className="close-btn">×</button>
        </div>
      )}
      <div className="workbook-container">
        <Workbook key={workbookKey} data={sheets} onChange={handleDataChange} />
      </div>
    </div>
  );
}
