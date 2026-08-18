import { useState, useCallback, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import { Upload, X, FileSpreadsheet, ChevronUp, ChevronDown, ArrowUpDown, Menu, Plus, RefreshCw, RotateCw, Trash2, Home, FilePlus } from 'lucide-react';
import './App.css';

// テンプレートタイプ定義
type TemplateType = 'sekou3' | 'normal3' | 'normal2';

const TEMPLATE_OPTIONS: { value: TemplateType; label: string; photosPerPage: number }[] = [
  { value: 'sekou3', label: '施工写真（3枚）', photosPerPage: 3 },
  { value: 'normal3', label: '3枚用', photosPerPage: 3 },
  { value: 'normal2', label: '2枚用', photosPerPage: 2 },
];

// 表示項目の定義
const DISPLAY_FIELD_OPTIONS = [
  { key: 'date', label: '日付' },
  { key: 'location', label: '場所' },
  { key: 'category', label: '種別' },
  { key: 'description', label: '内容' },
  { key: 'testDetails', label: '試験詳細' },
] as const;

type DisplayFieldKey = typeof DISPLAY_FIELD_OPTIONS[number]['key'];

const DEFAULT_DISPLAY_FIELDS: DisplayFieldKey[] = ['date', 'location', 'category', 'description'];

const TEST_TYPES = [
  { value: '', label: '選択なし' },
  { value: '水圧試験', label: '水圧試験' },
  { value: '耐圧試験', label: '耐圧試験' },
  { value: '通水試験', label: '通水試験' },
] as const;

const TEST_TEMPLATES: Record<string, { description: string; fields: { key: string; label: string; placeholder?: string; type?: 'select' | 'text'; options?: {value: string; label: string}[] }[] }> = {
  '水圧試験': {
    description: '水圧試験',
    fields: [
      { key: 'testPressure', label: '試験圧力', placeholder: '例: 1.0 Mpa' },
      { key: 'holdTime', label: '保持時間', placeholder: '例: 24h以上' },
      { key: 'startTime', label: '開始時間', placeholder: '例: 11:03' },
      { key: 'pressureState', label: '撮影対象', type: 'select', options: [{value: '', label: '選択してください'}, {value: '始圧', label: '始圧'}, {value: '終圧', label: '終圧'}] },
      { key: 'measuredPressure', label: '測定値', placeholder: '例: 1.0 Mpa' },
    ]
  },
  '耐圧試験': {
    description: '耐圧試験',
    fields: [
      { key: 'testPressure', label: '試験圧力', placeholder: '例: 1.75 Mpa' },
      { key: 'holdTime', label: '保持時間', placeholder: '例: 10分以上' },
      { key: 'startTime', label: '開始時間', placeholder: '例: 14:00' },
      { key: 'pressureState', label: '撮影対象', type: 'select', options: [{value: '', label: '選択してください'}, {value: '始圧', label: '始圧'}, {value: '終圧', label: '終圧'}] },
      { key: 'measuredPressure', label: '測定値', placeholder: '例: 1.75 Mpa' },
    ]
  },
  '通水試験': {
    description: '通水試験',
    fields: [
      { key: 'waterLocation', label: '注水場所', placeholder: '例: 5F系統' },
      { key: 'waterAmount', label: '注水量', placeholder: '例: 500L' },
      { key: 'waterStatus', label: '採水状況', placeholder: '例: 異常なし' },
    ]
  },
};

export interface PhotoData {
  id: string;
  file: File | null;
  previewUrl: string;
  date: string;
  location: string;
  category: string;
  testType: string;
  description: string;
  testFields: Record<string, string>;
  isBlank: boolean;
  rotation: number;
  displayFields: DisplayFieldKey[];
  locationNumber: string;
}

const CATEGORIES = [
  "配管工事",
  "ダクト工事",
  "保温工事",
  "塗装工事",
  "その他"
];

function chunkArray<T>(array: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
}

const GAS_URL = 'https://script.google.com/macros/s/AKfycbxjEZQdrAZ0tZN8EFNmZS2_FT9yabX4R2mmyjaPkmqmkDF9dvJikyRYkjXXZEW6jzJm/exec';

function IndividualDropzone({ onDropBlock }: { onDropBlock: (file: File) => void }) {
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop: (accepted) => {
      if (accepted.length > 0) onDropBlock(accepted[0]);
    },
    accept: { 'image/*': ['.jpeg', '.jpg', '.png', '.gif', '.webp'] },
    multiple: false
  });

  const rootProps = getRootProps();

  return (
    <div 
      {...rootProps} 
      className={`individual-dropzone ${isDragActive ? 'active' : ''}`} 
      onClick={(e) => { e.stopPropagation(); rootProps.onClick?.(e); }}
    >
      <input {...getInputProps()} />
      <Upload size={24} />
      <span>クリックまたはドロップ</span>
    </div>
  );
}

type ViewMode = 'home' | 'editor';

function App() {
  const [currentView, setCurrentView] = useState<ViewMode>('home');
  const [lastExportedSpreadsheetId, setLastExportedSpreadsheetId] = useState('');
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [photos, setPhotos] = useState<PhotoData[]>([]);
  const [projectNameLine1, setProjectNameLine1] = useState('');
  const [projectNameLine2, setProjectNameLine2] = useState('');
  const [selectedPhotoId, setSelectedPhotoId] = useState<string | null>(null);
  const [deleteDialogTarget, setDeleteDialogTarget] = useState<string | null>(null);
  const [moveDialogTarget, setMoveDialogTarget] = useState<string | null>(null);
  const [moveToPosition, setMoveToPosition] = useState('');
  
  // Template selection
  const [templateType, setTemplateType] = useState<TemplateType>('sekou3');
  
  // Global display fields
  const [globalDisplayFields, setGlobalDisplayFields] = useState<DisplayFieldKey[]>([...DEFAULT_DISPLAY_FIELDS]);
  
  // Import states
  const [fileList, setFileList] = useState<{id: string, name: string, dateStr: string}[]>([]);
  const [selectedFileId, setSelectedFileId] = useState('');
  const [isFetchingList, setIsFetchingList] = useState(false);
  
  const [isExporting, setIsExporting] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  
  // Mobile sidebar state
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  // Project suggestions
  const [buildingSuggestions, setBuildingSuggestions] = useState<string[]>([]);
  const [workSuggestions, setWorkSuggestions] = useState<Record<string, string[]>>({});

  // Get photos per page based on template
  const photosPerPage = TEMPLATE_OPTIONS.find(t => t.value === templateType)?.photosPerPage ?? 3;
  const hasLocationNumber = templateType === 'normal3' || templateType === 'normal2';

  const onDrop = useCallback((acceptedFiles: File[]) => {
    const newPhotos = acceptedFiles.map(file => ({
      id: crypto.randomUUID(),
      file,
      previewUrl: URL.createObjectURL(file),
      date: new Date().toISOString().split('T')[0],
      location: "",
      category: "",
      testType: "",
      description: "",
      testFields: {} as Record<string, string>,
      isBlank: false,
      rotation: 0,
      displayFields: [] as DisplayFieldKey[],
      locationNumber: "",
    }));
    setPhotos(prev => [...prev, ...newPhotos]);
  }, []);

  const { getRootProps, getInputProps, isDragActive, open } = useDropzone({
    onDrop,
    noClick: true,
    noKeyboard: true,
    accept: {
      'image/*': ['.jpeg', '.jpg', '.png', '.gif', '.webp']
    }
  });

  const fetchFileList = useCallback(async () => {
    setIsFetchingList(true);
    try {
      const res = await fetch(`${GAS_URL}?action=list`);
      const result = await res.json();
      if (result.success && result.files) {
        setFileList(result.files);
        setSelectedFileId('');
      } else {
        console.error('List error:', result.error);
      }
    } catch (err) {
      console.error('Failed to fetch file list:', err);
    } finally {
      setIsFetchingList(false);
    }
  }, []);

  const fetchProjectList = useCallback(async () => {
    try {
      const res = await fetch(`${GAS_URL}?action=listProjects`);
      const result = await res.json();
      if (result.success) {
        setBuildingSuggestions(result.buildings || []);
        setWorkSuggestions(result.works || {});
      }
    } catch (err) {
      console.error('Failed to fetch project list:', err);
    }
  }, []);

  useEffect(() => {
    fetchFileList();
    fetchProjectList();
  }, [fetchFileList, fetchProjectList]);

  const updatePhoto = (id: string, field: keyof PhotoData, value: string) => {
    setPhotos(prev => prev.map(p => {
      if (p.id !== id) return p;
      const updated = { ...p, [field]: value };
      if (field === 'testType') {
        const template = TEST_TEMPLATES[value];
        if (template) {
          updated.description = template.description;
          const newFields: Record<string, string> = {};
          template.fields.forEach(f => {
            newFields[f.key] = p.testFields[f.key] || '';
          });
          updated.testFields = newFields;
          // 試験区分を選択した場合、testDetailsを自動追加
          if (!globalDisplayFields.includes('testDetails')) {
            setGlobalDisplayFields(prev => [...prev, 'testDetails']);
          }
        } else {
          updated.description = '';
          updated.testFields = {};
        }
      }
      return updated;
    }));
  };

  const updateTestField = (id: string, fieldKey: string, value: string) => {
    setPhotos(prev => prev.map(p => {
      if (p.id !== id) return p;
      return { ...p, testFields: { ...p.testFields, [fieldKey]: value } };
    }));
  };

  const toggleGlobalDisplayField = (fieldKey: DisplayFieldKey) => {
    setGlobalDisplayFields(prev => 
      prev.includes(fieldKey)
        ? prev.filter(f => f !== fieldKey)
        : [...prev, fieldKey]
    );
  };

  // 試験区分を持つ写真があるかどうか
  const hasAnyTestType = photos.some(p => p.testType !== '');

  const showDeleteDialog = (id: string) => { setDeleteDialogTarget(id); };

  const deleteAndShift = () => {
    if (!deleteDialogTarget) return;
    setPhotos(prev => {
      const target = prev.find(p => p.id === deleteDialogTarget);
      if (target && target.previewUrl) URL.revokeObjectURL(target.previewUrl);
      return prev.filter(p => p.id !== deleteDialogTarget);
    });
    setDeleteDialogTarget(null);
    setSelectedPhotoId(null);
  };

  const deleteAndKeepBlank = () => {
    if (!deleteDialogTarget) return;
    setPhotos(prev => prev.map(p => {
      if (p.id !== deleteDialogTarget) return p;
      if (p.previewUrl) URL.revokeObjectURL(p.previewUrl);
      return { ...p, file: null, previewUrl: '', isBlank: true, rotation: 0 };
    }));
    setDeleteDialogTarget(null);
    setSelectedPhotoId(null);
  };

  const rotatePhoto = (id: string) => {
    setPhotos(prev => prev.map(p => {
      if (p.id !== id) return p;
      return { ...p, rotation: (p.rotation + 90) % 360 };
    }));
  };

  const replaceTargetPhoto = (id: string, newFile: File) => {
    setPhotos(prev => prev.map(p => {
      if (p.id !== id) return p;
      if (p.previewUrl) URL.revokeObjectURL(p.previewUrl);
      return {
        ...p,
        file: newFile,
        previewUrl: URL.createObjectURL(newFile),
        isBlank: false,
        rotation: 0
      };
    }));
  };

  const moveUp = (id: string) => {
    setPhotos(prev => {
      const idx = prev.findIndex(p => p.id === id);
      if (idx <= 0) return prev;
      const newArr = [...prev];
      [newArr[idx - 1], newArr[idx]] = [newArr[idx], newArr[idx - 1]];
      return newArr;
    });
  };

  const moveDown = (id: string) => {
    setPhotos(prev => {
      const idx = prev.findIndex(p => p.id === id);
      if (idx < 0 || idx >= prev.length - 1) return prev;
      const newArr = [...prev];
      [newArr[idx], newArr[idx + 1]] = [newArr[idx + 1], newArr[idx]];
      return newArr;
    });
  };

  const moveToIndex = (id: string) => {
    const targetPos = parseInt(moveToPosition);
    if (isNaN(targetPos) || targetPos < 1 || targetPos > photos.length) return;
    setPhotos(prev => {
      const fromIdx = prev.findIndex(p => p.id === id);
      if (fromIdx < 0) return prev;
      const toIdx = targetPos - 1;
      if (fromIdx === toIdx) return prev;
      const newArr = [...prev];
      const [item] = newArr.splice(fromIdx, 1);
      newArr.splice(toIdx, 0, item);
      return newArr;
    });
    setMoveDialogTarget(null);
    setMoveToPosition('');
  };

  const rotateAndCropImage = (file: File, rotation: number): Promise<string> => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const img = new Image();
        img.onload = () => {
          const tempCanvas = document.createElement('canvas');
          const tempCtx = tempCanvas.getContext('2d');
          if (!tempCtx) {
            resolve(reader.result as string);
            return;
          }
          
          const is90or270 = rotation === 90 || rotation === 270;
          tempCanvas.width = is90or270 ? img.height : img.width;
          tempCanvas.height = is90or270 ? img.width : img.height;
          
          tempCtx.translate(tempCanvas.width / 2, tempCanvas.height / 2);
          tempCtx.rotate((rotation * Math.PI) / 180);
          tempCtx.drawImage(img, -img.width / 2, -img.height / 2);

          const targetRatio = 4 / 3;
          const sWidth = tempCanvas.width;
          const sHeight = tempCanvas.height;
          const sRatio = sWidth / sHeight;
          
          let cropWidth = sWidth;
          let cropHeight = sHeight;
          let sx = 0;
          let sy = 0;

          if (sRatio > targetRatio) {
            cropWidth = sHeight * targetRatio;
            sx = (sWidth - cropWidth) / 2;
          } else {
            cropHeight = sWidth / targetRatio;
            sy = (sHeight - cropHeight) / 2;
          }

          const MAX_WIDTH = 1200;
          let dWidth = cropWidth;
          if (dWidth > MAX_WIDTH) dWidth = MAX_WIDTH;
          const dHeight = dWidth / targetRatio;

          const canvas = document.createElement('canvas');
          const ctx = canvas.getContext('2d');
          if (!ctx) return resolve(tempCanvas.toDataURL('image/jpeg', 0.85));

          canvas.width = dWidth;
          canvas.height = dHeight;
          ctx.drawImage(tempCanvas, sx, sy, cropWidth, cropHeight, 0, 0, dWidth, dHeight);
          
          resolve(canvas.toDataURL('image/jpeg', 0.85));
        };
        img.src = reader.result as string;
      };
      reader.readAsDataURL(file);
    });
  };

  const handleExport = async () => {
    if (photos.length === 0) {
      alert("エクスポートする写真がありません。");
      return;
    }
    setIsExporting(true);
    try {
      const photosData = await Promise.all(photos.map(async (p) => {
        let imageBase64 = '';
        if (p.file && !p.isBlank) {
          imageBase64 = await rotateAndCropImage(p.file, p.rotation || 0);
        }
        return { ...p, file: undefined, previewUrl: undefined, imageBase64, displayFields: [...globalDisplayFields] };
      }));

      const payload = {
        action: 'export',
        templateType,
        projectNameLine1,
        projectNameLine2,
        photos: photosData,
      };

      const res = await fetch(GAS_URL, {
        method: 'POST',
        body: JSON.stringify(payload),
      });
      const result = await res.json();
      
      if (result.url) {
        // URLからIDを抽出
        const match = result.url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (match) setLastExportedSpreadsheetId(match[1]);
        // result.spreadsheetId があればそちらを使用
        if (result.spreadsheetId) setLastExportedSpreadsheetId(result.spreadsheetId);
        alert(`保存しました！`);
        // window.openは削除しない
        window.open(result.url, '_blank');
      } else {
        alert('エクスポートに失敗しました: ' + (result.error || '不明なエラー'));
      }
    } catch (err: any) {
      const msg = err?.message || String(err);
      alert('エクスポート中にエラーが発生しました:\n' + msg);
      console.error(err);
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportPdf = async () => {
    if (!lastExportedSpreadsheetId) {
      alert('先に保存してからPDF出力してください');
      return;
    }
    setIsExportingPdf(true);
    try {
      const res = await fetch(GAS_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'exportPdf', spreadsheetId: lastExportedSpreadsheetId }),
      });
      const result = await res.json();
      if (result.pdfUrl) {
        window.open(result.pdfUrl, '_blank');
      } else {
        alert('PDF出力に失敗しました: ' + (result.error || '不明なエラー'));
      }
    } catch (err: any) {
      alert('PDF出力中にエラー:\n' + (err?.message || String(err)));
    } finally {
      setIsExportingPdf(false);
    }
  };

  const startNewLedger = () => {
    setPhotos([]);
    setProjectNameLine1('');
    setProjectNameLine2('');
    setTemplateType('sekou3');
    setGlobalDisplayFields([...DEFAULT_DISPLAY_FIELDS]);
    setSelectedPhotoId(null);
    setLastExportedSpreadsheetId('');
    setCurrentView('editor');
  };

  const goHome = () => {
    setCurrentView('home');
  };

  const handleImport = async () => {
    if (!selectedFileId) return;
    setIsImporting(true);
    try {
      const res = await fetch(`${GAS_URL}?action=import&spreadsheetId=${selectedFileId}`);
      const result = await res.json();
      
      if (result.data) {
        setProjectNameLine1(result.data.projectNameLine1 || '');
        setProjectNameLine2(result.data.projectNameLine2 || '');
        if (result.data.templateType) {
          setTemplateType(result.data.templateType);
        }
        // グローバル表示項目を復元（最初の写真のdisplayFieldsから）
        if (result.data.photos?.[0]?.displayFields) {
          setGlobalDisplayFields(result.data.photos[0].displayFields);
        }
        const restoredPhotos: PhotoData[] = (result.data.photos || []).map((p: any) => ({
          ...p,
          id: crypto.randomUUID(),
          file: null,
          previewUrl: p.imageUrl || '',
          isBlank: !p.imageUrl,
          testFields: p.testFields || {},
          rotation: 0,
          displayFields: p.displayFields || [...DEFAULT_DISPLAY_FIELDS],
          locationNumber: p.locationNumber || '',
        }));
        setPhotos(restoredPhotos);
        alert('データを読み込みました');
        setIsSidebarOpen(false);
      } else {
        alert(result.error || '読み込みに失敗しました');
      }
    } catch (err) {
      alert('インポート中にエラーが発生しました');
      console.error(err);
    } finally {
      setIsImporting(false);
    }
  };

  const photoPages = chunkArray(photos, photosPerPage);
  const totalPhotos = photos.length;

  if (currentView === 'home') {
    return (
      <div className="home-page">
        <div className="home-header">
          <h1>工事写真台帳</h1>
          <p>パル設計</p>
        </div>
        <div className="home-content">
          <button className="btn btn-primary btn-large" onClick={startNewLedger}>
            <FilePlus size={24} />
            新規作成
          </button>
          
          <div className="home-section">
            <div className="file-list-header">
              <h2>既存の台帳を編集</h2>
              <button 
                className="btn-icon" 
                onClick={fetchFileList} 
                disabled={isFetchingList}
                title="リストを更新"
              >
                <RefreshCw size={16} className={isFetchingList ? 'spinning' : ''} />
              </button>
            </div>
            <div className="home-file-list">
              {isFetchingList ? (
                <div className="file-list-empty">読み込み中...</div>
              ) : fileList.length === 0 ? (
                <div className="file-list-empty">台帳がありません</div>
              ) : (
                fileList.map((file) => (
                  <div key={file.id} className="home-file-item" onClick={async () => {
                    setSelectedFileId(file.id);
                    setIsImporting(true);
                    try {
                      const res = await fetch(`${GAS_URL}?action=import&spreadsheetId=${file.id}`);
                      const result = await res.json();
                      if (result.data) {
                        setProjectNameLine1(result.data.projectNameLine1 || '');
                        setProjectNameLine2(result.data.projectNameLine2 || '');
                        if (result.data.templateType) setTemplateType(result.data.templateType);
                        if (result.data.photos?.[0]?.displayFields) setGlobalDisplayFields(result.data.photos[0].displayFields);
                        const restoredPhotos: PhotoData[] = (result.data.photos || []).map((p: any) => ({
                          ...p,
                          id: crypto.randomUUID(),
                          file: null,
                          previewUrl: p.imageUrl || '',
                          isBlank: !p.imageUrl,
                          testFields: p.testFields || {},
                          rotation: 0,
                          displayFields: p.displayFields || [...DEFAULT_DISPLAY_FIELDS],
                          locationNumber: p.locationNumber || '',
                        }));
                        setPhotos(restoredPhotos);
                        setCurrentView('editor');
                      } else {
                        alert(result.error || '読み込みに失敗しました');
                      }
                    } catch (err) {
                      alert('読み込み中にエラーが発生しました');
                    } finally {
                      setIsImporting(false);
                    }
                  }}>
                    <FileSpreadsheet size={20} />
                    <div className="home-file-info">
                      <span className="home-file-name">{file.name}</span>
                      <span className="home-file-date">{file.dateStr}</span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="app-layout">
      {/* Mobile Drawer Overlay */}
      {isSidebarOpen && <div className="sidebar-overlay" onClick={() => setIsSidebarOpen(false)} />}
      
      {/* Left Sidebar */}
      <aside className={`sidebar ${isSidebarOpen ? 'is-open' : ''}`}>
        <div className="sidebar-header">
          <h1>工事写真台帳</h1>
        </div>
        
        <div className="sidebar-content">
          {/* テンプレート選択 */}
          <div className="sidebar-section">
            <h3>テンプレート</h3>
            <div className="template-selector">
              {TEMPLATE_OPTIONS.map(opt => (
                <label key={opt.value} className={`template-option ${templateType === opt.value ? 'selected' : ''}`}>
                  <input
                    type="radio"
                    name="template"
                    value={opt.value}
                    checked={templateType === opt.value}
                    onChange={() => setTemplateType(opt.value)}
                  />
                  <span className="template-option-label">{opt.label}</span>
                  <span className="template-option-count">{opt.photosPerPage}枚/ページ</span>
                </label>
              ))}
            </div>
          </div>

          {/* 表示項目チェックボックス */}
          <div className="sidebar-section">
            <h3>表示項目</h3>
            <div className="display-fields-section">
              <div className="display-fields-checkboxes">
                {DISPLAY_FIELD_OPTIONS.map(opt => {
                  if (opt.key === 'testDetails' && !hasAnyTestType) return null;
                  return (
                    <label key={opt.key} className="display-field-checkbox">
                      <input
                        type="checkbox"
                        checked={globalDisplayFields.includes(opt.key)}
                        onChange={() => toggleGlobalDisplayField(opt.key)}
                      />
                      <span>{opt.label}</span>
                    </label>
                  );
                })}
              </div>
            </div>
          </div>

          <div className="sidebar-section">
            <h3>工事名称</h3>
            <div className="form-group">
              <label>建物名称</label>
              <input
                type="text"
                list="building-suggestions"
                value={projectNameLine1}
                onChange={(e) => setProjectNameLine1(e.target.value)}
                placeholder="建物名称を入力..."
              />
              <datalist id="building-suggestions">
                {buildingSuggestions.map(b => (
                  <option key={b} value={b} />
                ))}
              </datalist>
            </div>
            <div className="form-group">
              <label>工事内容</label>
              <input
                type="text"
                list="work-suggestions"
                value={projectNameLine2}
                onChange={(e) => setProjectNameLine2(e.target.value)}
                placeholder="工事内容を入力..."
              />
              <datalist id="work-suggestions">
                {(workSuggestions[projectNameLine1] || []).map(w => (
                  <option key={w} value={w} />
                ))}
              </datalist>
            </div>
          </div>

        </div>
      </aside>

      {/* Main Content Area (Dropzone applied globally here) */}
      <main className="main-wrapper" {...getRootProps()}>
        <input {...getInputProps()} />
        
        {/* Drop Highlight Overlay */}
        <div className={`dropzone-overlay ${isDragActive ? 'is-active' : ''}`}>
          <Upload size={64} className="drop-icon-large" />
          <h2>写真をドロップして追加</h2>
        </div>

        {/* Top Navbar */}
        <header className="topbar">
          <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
            <button className="btn-icon" onClick={goHome} title="ホームに戻る">
              <Home size={24} />
            </button>
            <button className="menu-toggle" onClick={() => setIsSidebarOpen(true)}>
              <Menu size={24} />
            </button>
            <span className="topbar-title">台帳プレビュー</span>
          </div>
          
          <div className="topbar-actions">
            <button className="btn btn-primary" onClick={open}>
              <Plus size={18} />
              選択して追加
            </button>
            <button 
              className="btn btn-primary" 
              onClick={handleExport}
              disabled={isExporting || totalPhotos === 0}
              style={{ background: '#000' }}
            >
              <FileSpreadsheet size={18} />
              {isExporting ? '保存中...' : '保存'}
            </button>
            <button 
              className="btn btn-primary" 
              onClick={handleExportPdf}
              disabled={isExportingPdf || !lastExportedSpreadsheetId}
              style={{ background: '#c62828' }}
            >
              {isExportingPdf ? 'PDF出力中...' : 'PDF出力'}
            </button>
          </div>
        </header>

        {/* Scrollable Content */}
        <div className="main-content">
          {totalPhotos === 0 ? (
            <div className="empty-state">
              <Upload size={48} style={{ color: "var(--sys-text-muted)", marginBottom: "1rem" }} />
              <h2>写真がありません</h2>
              <p>右上の「選択して追加」ボタンをクリックするか、<br/>この画面のどこにでも写真をドラッグ＆ドロップしてください。</p>
            </div>
          ) : (
            <>
            <div className="page-container">
              {photoPages.map((pagePhotos, pageIndex) => (
                <div key={pageIndex} className="photo-page">
                  <div className="page-header">
                    <span>{pageIndex + 1} / {photoPages.length} ページ</span>
                    <button 
                      className="btn btn-danger" 
                      style={{ padding: '0.3rem 0.6rem', fontSize: '0.8rem' }}
                      onClick={() => {
                        if (confirm(`${pageIndex + 1}ページの写真をすべて削除しますか？`)) {
                          const idsToRemove = pagePhotos.map(p => p.id);
                          setPhotos(prev => {
                            prev.forEach(p => { if (idsToRemove.includes(p.id) && p.previewUrl && p.file) URL.revokeObjectURL(p.previewUrl); });
                            return prev.filter(p => !idsToRemove.includes(p.id));
                          });
                          setSelectedPhotoId(null);
                        }
                      }}
                    >
                      <Trash2 size={14} />
                      ページ削除
                    </button>
                  </div>
                  
                  {pagePhotos.map((photo) => {
                    const globalIndex = photos.findIndex(p => p.id === photo.id);
                    const template = TEST_TEMPLATES[photo.testType];
                    const isSelected = selectedPhotoId === photo.id;

                    return (
                      <div 
                        key={photo.id} 
                        className={`photo-item ${isSelected ? 'photo-item-selected' : ''} ${photo.isBlank ? 'photo-item-blank' : ''}`}
                        onClick={() => setSelectedPhotoId(isSelected ? null : photo.id)}
                      >
                        <div className="photo-number-badge">
                          {globalIndex + 1}/{totalPhotos}
                        </div>

                        {isSelected && (
                          <div className="photo-controls" onClick={(e) => e.stopPropagation()}>
                            <button 
                              className="ctrl-btn" 
                              onClick={() => moveUp(photo.id)}
                              disabled={globalIndex === 0}
                              title="上へ移動"
                            >
                              <ChevronUp size={16} />
                            </button>
                            <button 
                              className="ctrl-btn" 
                              onClick={() => moveDown(photo.id)}
                              disabled={globalIndex === totalPhotos - 1}
                              title="下へ移動"
                            >
                              <ChevronDown size={16} />
                            </button>
                            <button 
                              className="ctrl-btn" 
                              onClick={() => { setMoveDialogTarget(photo.id); setMoveToPosition(String(globalIndex + 1)); }}
                              title="指定位置へ移動"
                            >
                              <ArrowUpDown size={14} />
                            </button>
                          </div>
                        )}

                        <div className="photo-preview">
                          {!photo.isBlank && (
                            <>
                              <button 
                                className="rotate-btn" 
                                onClick={(e) => { e.stopPropagation(); rotatePhoto(photo.id); }}
                                title="回転"
                              >
                                <RotateCw size={14} />
                              </button>
                              <button 
                                className="delete-btn" 
                                onClick={(e) => { e.stopPropagation(); showDeleteDialog(photo.id); }}
                              >
                                <X size={14} />
                              </button>
                            </>
                          )}
                          {photo.isBlank ? (
                            <IndividualDropzone onDropBlock={(f) => replaceTargetPhoto(photo.id, f)} />
                          ) : (
                            <img src={photo.previewUrl} alt="プレビュー" style={{ transform: `rotate(${photo.rotation || 0}deg)` }} />
                          )}
                        </div>
                        
                        <div className="photo-info" onClick={(e) => e.stopPropagation()}>
                          {globalDisplayFields.includes('date') && (
                            <div className="info-row">
                              <label>日付</label>
                              <div className="input-wrapper" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                                <input 
                                  type="date" 
                                  value={photo.date} 
                                  onChange={(e) => updatePhoto(photo.id, 'date', e.target.value)} 
                                  style={{ flex: 1 }}
                                />
                                <button 
                                  type="button"
                                  className="btn-icon" 
                                  onClick={(e) => { e.stopPropagation(); updatePhoto(photo.id, 'date', ''); }}
                                  title="日付をクリア"
                                  style={{ padding: '0.5rem', background: 'var(--sys-bg)', border: '1px solid var(--sys-border)' }}
                                >
                                  <X size={16} />
                                </button>
                              </div>
                            </div>
                          )}
                          
                          {globalDisplayFields.includes('location') && (
                            <div className="info-row">
                              <label>場所</label>
                              <div className="input-wrapper">
                                <input 
                                  type="text"
                                  placeholder="場所を入力..."
                                  value={photo.location} 
                                  onChange={(e) => updatePhoto(photo.id, 'location', e.target.value)} 
                                />
                              </div>
                            </div>
                          )}
                          
                          {globalDisplayFields.includes('category') && (
                            <div className="info-row">
                              <label>種別</label>
                              <select 
                                value={photo.category} 
                                onChange={(e) => updatePhoto(photo.id, 'category', e.target.value)}
                              >
                                <option value="">選択してください</option>
                                {CATEGORIES.map(cat => (
                                  <option key={cat} value={cat}>{cat}</option>
                                ))}
                              </select>
                            </div>
                          )}

                          <div className="info-row">
                            <label>試験区分</label>
                            <select 
                              value={photo.testType} 
                              onChange={(e) => updatePhoto(photo.id, 'testType', e.target.value)}
                            >
                              {TEST_TYPES.map(t => (
                                <option key={t.value} value={t.value}>{t.label}</option>
                              ))}
                            </select>
                          </div>
                          
                          {globalDisplayFields.includes('description') && (
                            <div className="info-row">
                              <label>内容</label>
                              <div className="input-wrapper">
                                <textarea 
                                  placeholder="内容を入力..."
                                  rows={2}
                                  value={photo.description} 
                                  onChange={(e) => updatePhoto(photo.id, 'description', e.target.value)} 
                                />
                              </div>
                            </div>
                          )}

                          {template && globalDisplayFields.includes('testDetails') && (
                            <div className="test-fields-card">
                              <div className="test-fields-header">{photo.testType} 詳細</div>
                              {template.fields.map(f => (
                                <div key={f.key} className="info-row">
                                  <label>{f.label}</label>
                                  <div className="input-wrapper">
                                    {f.type === 'select' ? (
                                      <select
                                        value={photo.testFields[f.key] || ''}
                                        onChange={(e) => updateTestField(photo.id, f.key, e.target.value)}
                                        className="test-field-select"
                                      >
                                        {f.options?.map(opt => (
                                          <option key={opt.value} value={opt.value}>{opt.label}</option>
                                        ))}
                                      </select>
                                    ) : (
                                      <input 
                                        type="text" 
                                        placeholder={f.placeholder}
                                        value={photo.testFields[f.key] || ''} 
                                        onChange={(e) => updateTestField(photo.id, f.key, e.target.value)} 
                                      />
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}

                          {/* 撮影場所番号（2枚用・3枚用のみ） */}
                          {hasLocationNumber && (
                            <div className="info-row">
                              <label>撮影場所No.</label>
                              <div className="input-wrapper">
                                <input
                                  type="text"
                                  placeholder="撮影場所番号..."
                                  value={photo.locationNumber}
                                  onChange={(e) => updatePhoto(photo.id, 'locationNumber', e.target.value)}
                                />
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              ))}
            </div>

            {/* ページ追加ボタン */}
            <button 
              className="btn btn-secondary btn-full"
              style={{ maxWidth: '850px', margin: '0 auto', padding: '1rem' }}
              onClick={() => {
                const newBlanks: PhotoData[] = Array.from({ length: photosPerPage }, () => ({
                  id: crypto.randomUUID(),
                  file: null,
                  previewUrl: '',
                  date: new Date().toISOString().split('T')[0],
                  location: '',
                  category: '',
                  testType: '',
                  description: '',
                  testFields: {} as Record<string, string>,
                  isBlank: true,
                  rotation: 0,
                  displayFields: [] as DisplayFieldKey[],
                  locationNumber: '',
                }));
                setPhotos(prev => [...prev, ...newBlanks]);
              }}
            >
              <Plus size={18} />
              ページを追加
            </button>
            </>
          )}
        </div>
      </main>

      {/* Delete Modal */}
      {deleteDialogTarget && (
        <div className="modal-overlay" onClick={() => setDeleteDialogTarget(null)}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <h3>写真の削除</h3>
            <p>この写真をどのように処理しますか？</p>
            <div className="modal-buttons">
              <button className="btn btn-danger btn-full" onClick={deleteAndShift}>
                削除して上に詰める
              </button>
              <button className="btn btn-primary btn-full" onClick={deleteAndKeepBlank} style={{ marginTop: '0.5rem' }}>
                写真を入れ替える（空枠にする）
              </button>
            </div>
            <button className="btn-cancel" onClick={() => setDeleteDialogTarget(null)}>
              キャンセル
            </button>
          </div>
        </div>
      )}

      {/* Move Modal */}
      {moveDialogTarget && (
        <div className="modal-overlay" onClick={() => setMoveDialogTarget(null)}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <h3>写真の移動</h3>
            <p>何枚目に移動しますか？</p>
            <div className="move-input-row">
              <input
                type="number"
                min={1}
                max={totalPhotos}
                value={moveToPosition}
                onChange={(e) => setMoveToPosition(e.target.value)}
                autoFocus
              />
              <span style={{color: 'var(--sys-text-muted)'}}>/ {totalPhotos}</span>
            </div>
            <button className="btn btn-primary btn-full" onClick={() => moveToIndex(moveDialogTarget)}>
              移動する
            </button>
            <button className="btn-cancel" onClick={() => setMoveDialogTarget(null)}>
              キャンセル
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
