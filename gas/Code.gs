/**
 * 工事写真台帳 - Google Apps Script バックエンド
 * 
 * Required OAuth Scopes:
 * DriveApp: https://www.googleapis.com/auth/drive
 * SpreadsheetApp: https://www.googleapis.com/auth/spreadsheets
 * UrlFetchApp: https://www.googleapis.com/auth/script.external_request
 */

const ROOT_FOLDER_NAME = '工事写真台帳';

// フォルダがあれば取得、なければ作成
function getOrCreateFolder(parent, name) {
  const p = parent || DriveApp.getRootFolder();
  const folderName = (name && String(name).trim()) ? String(name).trim() : '未設定';
  const folders = p.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return p.createFolder(folderName);
}

// 初回権限一括承認用（GASエディタで「▶ 実行」）
function authorizeAll() {
  const root = DriveApp.getRootFolder();
  Logger.log('Drive権限 OK: ' + root.getName());
  UrlFetchApp.fetch('https://www.google.com');
  Logger.log('UrlFetch権限 OK');
  Logger.log('すべての権限の承認が完了しました！');
}

const TEMPLATES = {
  'sekou3': {
    id: '1j8Nhr_j_Fj7AKkP5Unrlb5ExbxZ7K--8ysoqweNQEXY',
    blockDataSize: 19,      // 各ブロックのデータ行数
    blocksPerPage: 3,
    firstDataRow: 5,         // 最初のブロックのデータ開始行
    contentRows: 19,         // C列の書き込み可能行数
    hasLocationNumber: false,
    hasBlockHeader: false,   // ブロックごとのヘッダーなし
  },
  'normal3': {
    id: '1os1_yXlWJJGYnIh7SoI_yfzLJL0HVvC9qy26-prUdr0',
    blockDataSize: 21,       // ヘッダー1行 + データ20行
    blocksPerPage: 3,
    firstDataRow: 5,
    contentRows: 18,         // C列の書き込み可能行数（C5:C22の18行）
    locationNumberOffset: 18, // データ開始から18行目
    hasLocationNumber: true,
    hasBlockHeader: true,
  },
  'normal2': {
    id: '1bJkLH0U5IJL0LdJYztZ011thsGx9N62hKSfg8k22NTI',
    blockDataSize: 21,
    blocksPerPage: 2,
    firstDataRow: 5,
    contentRows: 18,
    locationNumberOffset: 18,
    hasLocationNumber: true,
    hasBlockHeader: true,
  },
};

function formatJapaneseDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'import') {
    const spreadsheetId = e.parameter.spreadsheetId;
    return handleImport(spreadsheetId);
  } else if (action === 'list') {
    return handleList();
  } else if (action === 'listProjects') {
    return handleListProjects();
  }
  
  return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.action === 'export') {
      return handleExport(data);
    } else if (data.action === 'exportPdf') {
      return handleExportPdf(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleExport(data) {
  const { projectNameLine1, projectNameLine2, photos, templateType = 'normal3', spreadsheetId } = data;
  
  const template = TEMPLATES[templateType];
  if (!template) {
    throw new Error('Invalid templateType: ' + templateType);
  }

  let ss;
  let targetFile;
  let mainFolder;
  let photosFolder;

  // 既存ファイルがある場合、テンプレートタイプ（2枚用 ↔ 3枚用 ↔ 施工写真用）が途中で変更されたかチェック
  let isTemplateChanged = false;
  if (spreadsheetId) {
    try {
      const checkSs = SpreadsheetApp.openById(spreadsheetId);
      const metaSheet = checkSs.getSheetByName('_metadata');
      if (metaSheet) {
        const jsonStr = metaSheet.getRange('A1').getValue();
        if (jsonStr) {
          const meta = JSON.parse(jsonStr);
          if (meta.templateType && meta.templateType !== templateType) {
            isTemplateChanged = true;
          }
        }
      }
    } catch (e) {
      Logger.log('Meta check error: ' + e.message);
    }
  }

  // テンプレートタイプが変わっていない場合のみ既存ファイルへ上書き保存。
  // タイプが「3枚用」等へ切り替えられた場合は、選択された新しいテンプレートから正しく台帳を生成！
  if (spreadsheetId && !isTemplateChanged) {
    try {
      targetFile = DriveApp.getFileById(spreadsheetId);
      ss = SpreadsheetApp.openById(spreadsheetId);
      const parents = targetFile.getParents();
      mainFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
      
      const rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
      const targetBuildingFolder = getOrCreateFolder(rootFolder, projectNameLine1 || '未設定');
      const targetWorkFolder = getOrCreateFolder(targetBuildingFolder, projectNameLine2 || '未設定');
      
      const currentWorkFolder = mainFolder.getParents().hasNext() ? mainFolder.getParents().next() : null;
      if (currentWorkFolder && currentWorkFolder.getId() !== targetWorkFolder.getId()) {
        mainFolder.moveTo(targetWorkFolder);
      }
      
      const photoFolders = mainFolder.getFoldersByName('写真');
      photosFolder = photoFolders.hasNext() ? photoFolders.next() : mainFolder.createFolder('写真');
    } catch (e) {
      Logger.log('Spreadsheet open/move error: ' + e.message);
    }
  }

  // IDがない、またはテンプレートタイプが変更された場合は、選択されたテンプレート（'normal3', 'sekou3', 'normal2'）から新しく生成！
  if (!ss) {
    const todayStr = new Date().toLocaleDateString('ja-JP').replace(/\//g, '');
    const exportFolderName = `工事写真台帳_${todayStr}`;
    
    const rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
    const buildingFolder = getOrCreateFolder(rootFolder, projectNameLine1 || '未設定');
    const workFolder = getOrCreateFolder(buildingFolder, projectNameLine2 || '未設定');
    mainFolder = workFolder.createFolder(exportFolderName);
    photosFolder = mainFolder.createFolder('写真');
    
    // 選択されたテンプレート（3枚用:1os1..., 施工写真:1j8N..., 2枚用:1bJk...）のファイルを複製
    const templateFile = DriveApp.getFileById(template.id);
    targetFile = templateFile.makeCopy(exportFolderName, mainFolder);
    ss = SpreadsheetApp.openById(targetFile.getId());
  }

  const sheet = ss.getSheets()[0];
  const targetId = ss.getId();
  
  const totalPhotos = photos.length;
  const blocksPerPage = template.blocksPerPage;
  const pageCount = Math.max(1, Math.ceil(totalPhotos / blocksPerPage));
  
  const templateSheet = ss.getSheets()[0];
  const sheets = [templateSheet];
  try {
    templateSheet.setName('ページ1');
  } catch (e) {}
  
  // 不足しているページ数ぶん、templateSheet を複製して「ページ2」「ページ3」... を100%確実に作成！
  for (let p = 1; p < pageCount; p++) {
    const sheetName = `ページ${p + 1}`;
    let sheetForPage = ss.getSheetByName(sheetName);
    if (!sheetForPage) {
      sheetForPage = templateSheet.copyTo(ss).setName(sheetName);
    }
    sheets.push(sheetForPage);
  }
  
  // 不要になった過去の余分なページタブがあれば自動削除整理
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const sh = allSheets[i];
    const sName = sh.getName();
    if (sName.startsWith('ページ')) {
      const pNum = parseInt(sName.replace('ページ', ''), 10);
      if (pNum > pageCount) {
        try {
          ss.deleteSheet(sh);
        } catch (e) {
          Logger.log('Delete sheet error: ' + e.message);
        }
      }
    }
  }
  
  // 各シートのヘッダー（A1:C1 = 建物名称, A2:C2 = 工事内容）の書き込みおよび初期化
  sheets.forEach((sh, sheetIdx) => {
    // 建物名称（A1:C1 結合セル）
    const buildingRange = sh.getRange('A1:C1');
    buildingRange.merge();
    sh.getRange('A1').setValue(projectNameLine1 || '');
    
    // 工事内容（A2:C2 結合セル）
    const workRange = sh.getRange('A2:C2');
    workRange.merge();
    sh.getRange('A2').setValue(projectNameLine2 || '');
    
    // 2枚目以降の複製タブの場合、古いデータ内容をクリア
    if (sheetIdx > 0) {
      for (let b = 0; b < blocksPerPage; b++) {
        const dataStartRow = template.firstDataRow + b * template.blockDataSize;
        sh.getRange(`A${dataStartRow}`).clearContent();
        sh.getRange(dataStartRow, 3, template.contentRows, 1).clearContent();
        if (template.hasLocationNumber) {
          sh.getRange(dataStartRow + template.locationNumberOffset, 3, 2, 1).clearContent();
        }
      }
    }
  });
  
  // 写真URLを記録する配列
  const photoImageUrls = new Array(totalPhotos).fill('');
  
  for (let i = 0; i < totalPhotos; i++) {
    const photo = photos[i];
    if (photo.isBlank) continue;
    
    const p = Math.floor(i / blocksPerPage);
    const b = i % blocksPerPage;
    const targetSheet = sheets[p];
    
    const dataStartRow = template.firstDataRow + b * template.blockDataSize;
    
    // 内容書き込み (C列に1行ずつ順に詰めて隙間なく配置)
    const fieldsToWrite = [];
    const displayFields = photo.displayFields || [];
    
    if (displayFields.includes('date') && photo.date) {
      fieldsToWrite.push(`日付： ${formatJapaneseDate(photo.date)}`);
    }
    if (displayFields.includes('location') && photo.location) {
      fieldsToWrite.push(`場所： ${photo.location}`);
    }
    if (displayFields.includes('category') && photo.category) {
      fieldsToWrite.push(`種別： ${photo.category}`);
    }
    if (displayFields.includes('description') && photo.description) {
      const descLines = String(photo.description).split('\n');
      descLines.forEach((line, lIdx) => {
        if (lIdx === 0) {
          fieldsToWrite.push(`内容： ${line}`);
        } else {
          fieldsToWrite.push(`　　　 ${line}`);
        }
      });
    }
    
    const hasTestType = !!photo.testType && displayFields.includes('testDetails');
    if (hasTestType) {
      // 試験記録の全行を準備（定義された順番通りに確実に取得）
      const testLines = ['【試験記録】'];
      if (photo.testFields) {
        // キーの書き込み順番を絶対固定（開始時間含む）
        const keyOrder = ['designPressure', 'holdTime', 'pressureState', 'startTime', 'testPressure', 'waterLocation', 'waterAmount', 'waterStatus'];
        keyOrder.forEach(key => {
          if (photo.testFields[key]) {
            testLines.push(`${getFieldLabel(key)}： ${photo.testFields[key]}`);
          }
        });
      }
      
      // 内容の有無に関わらず直後に上に詰めて【試験記録】を追加！
      fieldsToWrite.push(...testLines);
    }
    
    // 全体を最大18行枠内に上から順番に隙間なくきっちり書き込み！
    for (let j = 0; j < fieldsToWrite.length && j < template.contentRows; j++) {
      targetSheet.getRange(dataStartRow + j, 3).setValue(fieldsToWrite[j]);
    }
    
    // 撮影場所番号の書き込み
    if (template.hasLocationNumber && photo.locationNumber) {
      targetSheet.getRange(dataStartRow + template.locationNumberOffset, 3).setValue(photo.locationNumber);
    }
    
    // 写真の挿入 (該当タブシートの対象ブロックへ)
    if (photo.imageBase64) {
      try {
        const imageBlob = base64ToBlob(photo.imageBase64, `photo_${i + 1}`);
        const imageUrl = insertImageIntoCell(targetSheet, dataStartRow, template, imageBlob, photosFolder.getId());
        photoImageUrls[i] = imageUrl || '';
      } catch (err) {
        Logger.log('Image insertion error: ' + err.message);
      }
    } else if (photo.imageUrl) {
      photoImageUrls[i] = photo.imageUrl;
      // 既存画像の場合も写真結合セル(A:B)に=IMAGE数式を確実に配置
      const match = photo.imageUrl.match(/id=([a-zA-Z0-9_-]+)/);
      if (match) {
        const fileId = match[1];
        const lh3Url = `https://lh3.googleusercontent.com/d/${fileId}`;
        let rowSpan = 20;
        if (template && template.contentRows === 19) rowSpan = 19;
        const endRow = dataStartRow + rowSpan - 1;
        const photoRange = targetSheet.getRange(`A${dataStartRow}:B${endRow}`);
        photoRange.merge();
        targetSheet.getRange(`A${dataStartRow}`).setFormula(`=IMAGE("${lh3Url}", 2)`);
      }
    }
  }
  
  // 過去のメタデータから「前回保存されていたDrive画像ファイルID」のリストを保持
  const previousFileIds = [];
  const oldMetaSheet = ss.getSheetByName('_metadata');
  if (oldMetaSheet) {
    try {
      const jsonStr = oldMetaSheet.getRange('A1').getValue();
      if (jsonStr) {
        const meta = JSON.parse(jsonStr);
        (meta.photos || []).forEach(p => {
          const url = p.imageUrl || '';
          const match = url.match(/id=([a-zA-Z0-9_-]+)/);
          if (match) previousFileIds.push(match[1]);
        });
      }
    } catch (e) {
      Logger.log('Previous meta parse error: ' + e.message);
    }
  }

  // 今回の保存に残った画像ファイルIDのリスト
  const currentFileIds = photoImageUrls.map(url => {
    const match = (url || '').match(/id=([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }).filter(Boolean);

  // 前回存在したが、ページ削除等で今回不要になった画像ファイルをGoogle Drive上から自動でゴミ箱へ削除！
  previousFileIds.forEach(oldId => {
    if (!currentFileIds.includes(oldId)) {
      try {
        DriveApp.getFileById(oldId).setTrashed(true);
      } catch (e) {
        Logger.log('Failed to trash unused image file ' + oldId + ': ' + e.message);
      }
    }
  });

  // _metadataシートにJSONデータを保存
  let metaSheet = ss.getSheetByName('_metadata');
  if (!metaSheet) {
    metaSheet = ss.insertSheet('_metadata');
  }
  
  const metaData = {
    projectNameLine1,
    projectNameLine2,
    templateType,
    photos: photos.map((p, idx) => ({
      date: p.date,
      location: p.location,
      locationNumber: p.locationNumber,
      category: p.category,
      testType: p.testType,
      description: p.description,
      testFields: p.testFields,
      displayFields: p.displayFields,
      isBlank: p.isBlank,
      imageUrl: photoImageUrls[idx] || p.imageUrl || '',
    })),
    exportDate: new Date().toISOString(),
  };
  
  metaSheet.getRange('A1').setValue(JSON.stringify(metaData));
  metaSheet.hideSheet();
  
  const url = ss.getUrl();
  
  return ContentService.createTextOutput(JSON.stringify({ 
    success: true, 
    url: url,
    spreadsheetId: targetId
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleImport(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let metaSheet = ss.getSheetByName('_metadata');
    
    if (!metaSheet) {
      return ContentService.createTextOutput(JSON.stringify({ 
        error: 'メタデータが見つかりません。このスプレッドシートはツールからエクスポートされたものではありません。' 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const jsonStr = metaSheet.getRange('A1').getValue();
    const data = JSON.parse(jsonStr);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true, 
      data: data 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleList() {
  try {
    const templateIds = Object.values(TEMPLATES).map(t => t.id);
    const query = "title contains '工事写真台帳' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false";
    const files = DriveApp.searchFiles(query);
    const fileList = [];
    
    let count = 0;
    while (files.hasNext() && count < 30) {
      const file = files.next();
      const fileId = file.getId();
      
      // テンプレート用スプレッドシートは除外
      if (templateIds.includes(fileId)) continue;
      
      let buildingName = '';
      let workName = '';
      
      try {
        const ss = SpreadsheetApp.openById(fileId);
        const metaSheet = ss.getSheetByName('_metadata');
        if (metaSheet) {
          const jsonStr = metaSheet.getRange('A1').getValue();
          if (jsonStr) {
            const meta = JSON.parse(jsonStr);
            buildingName = meta.projectNameLine1 || '';
            workName = meta.projectNameLine2 || '';
          }
        }
      } catch (e) {
        // メタデータがない・読み込めない場合はスルー
      }
      
      fileList.push({
        id: fileId,
        name: file.getName(),
        buildingName: buildingName,
        workName: workName,
        date: file.getDateCreated().getTime()
      });
      count++;
    }
    
    fileList.sort((a, b) => b.date - a.date);
    
    const formattedList = fileList.map(item => {
      const d = new Date(item.date);
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      return {
        id: item.id,
        name: item.name,
        buildingName: item.buildingName,
        workName: item.workName,
        dateStr: dateStr
      };
    });

    return ContentService.createTextOutput(JSON.stringify({ 
      success: true, 
      files: formattedList 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleExportPdf(data) {
  try {
    const spreadsheetId = data.spreadsheetId;
    const file = DriveApp.getFileById(spreadsheetId);
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const parents = file.getParents();
    const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    
    // スプレッドシート本体から直接ネイティブPDF変換（URLパラメータによる変な幅合わせが100%発生しない公式機能）
    const pdfBlob = ss.getAs('application/pdf').setName(file.getName() + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      pdfUrl: pdfFile.getUrl()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleListProjects() {
  try {
    const rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (!rootFolders.hasNext()) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, buildings: [], works: {} 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const rootFolder = rootFolders.next();
    const buildings = [];
    const works = {};
    
    const buildingFolders = rootFolder.getFolders();
    while (buildingFolders.hasNext()) {
      const bf = buildingFolders.next();
      const bName = bf.getName();
      if (bName === '未設定') continue;
      buildings.push(bName);
      works[bName] = [];
      const workFolders = bf.getFolders();
      while (workFolders.hasNext()) {
        const wf = workFolders.next();
        const wName = wf.getName();
        if (wName !== '未設定') {
          works[bName].push(wName);
        }
      }
    }
    
    buildings.sort();
    
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true, buildings, works 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function base64ToBlob(base64Data, fileName) {
  const parts = base64Data.split(',');
  const contentType = parts[0].match(/:(.*?);/)[1];
  const raw = Utilities.base64Decode(parts[1]);
  // MIMEタイプから拡張子を決定
  const extMap = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/gif': '.gif', 'image/webp': '.webp' };
  const ext = extMap[contentType] || '.jpg';
  return Utilities.newBlob(raw, contentType, fileName + ext);
}

function insertImageIntoCell(sheet, startRow, template, imageBlob, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(imageBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  const fileId = file.getId();
  const cdnUrl = `https://lh3.googleusercontent.com/d/${fileId}`;
  const displayUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1600`;
  
  // 写真結合セルの正確な行数 (normal3/normal2: 20行, sekou3: 19行)
  let rowSpan = 20;
  if (template && template.contentRows === 19) {
    rowSpan = 19;
  }
  const endRow = startRow + rowSpan - 1;
  
  // A列とB列の結合セル範囲(例: A5:B24 または A5:B23)をしっかりと結合
  const photoRange = sheet.getRange(`A${startRow}:B${endRow}`);
  photoRange.merge();
  
  // モード2 (=IMAGE(cdnUrl, 2)): 写真枠のサイズ全体ぴったりに即座に余白なしで表示
  const anchorCell = sheet.getRange(`A${startRow}`);
  anchorCell.setFormula(`=IMAGE("${cdnUrl}", 2)`);
  
  return displayUrl;
}

function getFieldLabel(key) {
  const labels = {
    'designPressure': '設計圧力',
    'testPressure': '試験圧力',
    'holdTime': '保持時間',
    'startTime': '開始時間',
    'pressureState': '撮影対象',
    'startPressure': '始圧',
    'waterLocation': '注水場所',
    'waterAmount': '注水量',
    'waterStatus': '採水状況',
  };
  return labels[key] || key;
}
