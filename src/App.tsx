import { useState, useCallback, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import { Upload, X, FileSpreadsheet, ChevronUp, ChevronDown, ArrowUpDown, Menu, Plus, RefreshCw, RotateCw } from 'lucide-react';
import './App.css';

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

const GAS_URL = 'https://script.google.com/macros/s/AKfycbzYK9nQN6YFBbryRSXKJbbj8bCU-CnF_FwpnUhZ1U2_MVztsWYSq9w2d5kGDnNZhwwC/exec';

function IndividualDropzone({ onDropBlock }: { onDropBlock: (file: File) => void }) {
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop: (accepted) => {
      if (accepted.length > 0) onDropBlock(accepted[0]);
    },
    accept: { 'image/*': ['.jpeg', '.jpg', '.png', '.gif', '.webp'] },
    multiple: false
  });

  return (
    <div {...getRootProps()} className={`individual-dropzone ${isDragActive ? 'active' : ''}`} onClick={(e) => e.stopPropagation()}>
      <input {...getInputProps()} />
      <Upload size={24} />
      <span>クリックまたはドロップ</span>
    </div>
  );
}

function App() {
  const [photos, setPhotos] = useState<PhotoData[]>([]);
  const [projectNameLine1, setProjectNameLine1] = useState('');
  const [projectNameLine2, setProjectNameLine2] = useState('');
  const [selectedPhotoId, setSelectedPhotoId] = useState<string | null>(null);
  const [deleteDialogTarget, setDeleteDialogTarget] = useState<string | null>(null);
  const [moveDialogTarget, setMoveDialogTarget] = useState<string | null>(null);
  const [moveToPosition, setMoveToPosition] = useState('');
  
  // Import states
  const [fileList, setFileList] = useState<{id: string, name: string, dateStr: string}[]>([]);
  const [selectedFileId, setSelectedFileId] = useState('');
  const [isFetchingList, setIsFetchingList] = useState(false);
  
  const [isExporting, setIsExporting] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  
  // Mobile sidebar state
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

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
    }));
    setPhotos(prev => [...prev, ...newPhotos]);
  }, []);

  const { getRootProps, getInputProps, isDragActive, open } = useDropzone({
    onDrop,
    noClick: true, // Prevent clicking anywhere triggering the file dialog
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
        // リスト取得後に選択状態をクリア
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

  useEffect(() => {
    fetchFileList();
  }, [fetchFileList]);

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
          // 回転用の一時Canvasを作成
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

          // 正しい向きになった画像をベースにクロップ処理
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
        return { ...p, file: undefined, previewUrl: undefined, imageBase64 };
      }));

      const payload = {
        action: 'export',
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
        alert(`スプレッドシートが作成されました！\n${result.url}`);
        window.open(result.url, '_blank');
      } else {
        alert('エクスポートに失敗しました: ' + (result.error || '不明なエラー'));
      }
    } catch (err) {
      alert('エクスポート中にエラーが発生しました');
      console.error(err);
    } finally {
      setIsExporting(false);
    }
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
        const restoredPhotos: PhotoData[] = (result.data.photos || []).map((p: any) => ({
          ...p,
          id: crypto.randomUUID(),
          file: null,
          previewUrl: '', // URLs from GAS not supported yet without separate drive fetching
          testFields: p.testFields || {},
          rotation: 0,
        }));
        setPhotos(restoredPhotos);
        alert('データを読み込みました');
        setIsSidebarOpen(false); // Close sidebar on mobile
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

  const photoPages = chunkArray(photos, 3);
  const totalPhotos = photos.length;

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
          <div className="sidebar-section">
            <h3>工事名称</h3>
            <div className="form-group">
              <label>建物名称</label>
              <input
                type="text"
                value={projectNameLine1}
                onChange={(e) => setProjectNameLine1(e.target.value)}
              />
            </div>
            <div className="form-group">
              <label>工事内容</label>
              <input
                type="text"
                value={projectNameLine2}
                onChange={(e) => setProjectNameLine2(e.target.value)}
              />
            </div>
          </div>

          <div className="sidebar-section">
            <div className="file-list-header">
              <h3>既存の台帳を編集</h3>
              <button 
                className="btn-icon" 
                onClick={fetchFileList} 
                disabled={isFetchingList}
                title="リストを更新"
              >
                <RefreshCw size={14} className={isFetchingList ? 'animate-spin' : ''} />
              </button>
            </div>
            
            <div className="file-list-container">
              <div className="file-list">
                {isFetchingList ? (
                  <div className="file-list-empty">読み込み中...</div>
                ) : fileList.length === 0 ? (
                  <div className="file-list-empty">
                    最近の台帳が見つかりません。<br/>右上のボタンで更新してください。
                  </div>
                ) : (
                  fileList.map((file) => (
                    <label key={file.id} className="file-item">
                      <input 
                        type="radio" 
                        name="import-file" 
                        value={file.id} 
                        checked={selectedFileId === file.id}
                        onChange={() => setSelectedFileId(file.id)}
                      />
                      <div className="file-item-info">
                        <span className="file-item-name" title={file.name}>{file.name}</span>
                        <span className="file-item-date">{file.dateStr}</span>
                      </div>
                    </label>
                  ))
                )}
              </div>
            </div>
            <button 
              className="btn btn-secondary btn-full" 
              onClick={handleImport}
              disabled={!selectedFileId || isImporting}
            >
              <FileSpreadsheet size={16} />
              {isImporting ? '読込中...' : '選択した台帳を読み込む'}
            </button>
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
              <Upload size={18} />
              {isExporting ? '出力中...' : '出力する'}
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
            <div className="page-container">
              {photoPages.map((pagePhotos, pageIndex) => (
                <div key={pageIndex} className="photo-page">
                  <div className="page-header">
                    <span>{pageIndex + 1} / {photoPages.length} ページ</span>
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
                          
                          <div className="info-row">
                            <label>場所</label>
                            <div className="input-wrapper">
                              <textarea 
                                className="multiline-input"
                                placeholder="場所を入力..."
                                rows={2}
                                value={photo.location} 
                                onChange={(e) => updatePhoto(photo.id, 'location', e.target.value)} 
                              />
                            </div>
                          </div>
                          
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

                          {template && (
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
                        </div>
                      </div>
                    );
                  })}
                </div>
              ))}
            </div>
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
