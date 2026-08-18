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
    throw new Error('Invalid templateType');
  }

  // ユーザー指定の純粋な元テンプレートスプレッドシート本体のIDのみを常に使用（過去の変形された複製ファイルIDを完全遮断）
  const targetId = template.id;
  const ss = SpreadsheetApp.openById(targetId);
  const sheet = ss.getSheets()[0];
  
  const targetFile = DriveApp.getFileById(targetId);
  const parents = targetFile.getParents();
  const mainFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const photoFolders = mainFolder.getFoldersByName('写真');
  const photosFolder = photoFolders.hasNext() ? photoFolders.next() : mainFolder.createFolder('写真');
  
  const totalPhotos = photos.length;
  const existingBlocks = template.blocksPerPage;
  
  if (totalPhotos > existingBlocks) {
    const block1StartRow = template.firstDataRow - (template.hasBlockHeader ? 1 : 0);
    const block1Range = sheet.getRange(block1StartRow, 1, template.blockDataSize, 4); // A:D
    
    for (let i = existingBlocks; i < totalPhotos; i++) {
      const dataStartRow = template.firstDataRow + i * template.blockDataSize;
      const blockStartRow = dataStartRow - (template.hasBlockHeader ? 1 : 0);
      
      sheet.insertRowsAfter(blockStartRow - 1, template.blockDataSize);
      const targetRange = sheet.getRange(blockStartRow, 1);
      block1Range.copyTo(targetRange);
      
      // コピーしたブロックの内容をクリア
      sheet.getRange(`A${dataStartRow}`).clearContent();
      sheet.getRange(dataStartRow, 3, template.contentRows, 1).clearContent();
      
      if (template.hasLocationNumber) {
        sheet.getRange(dataStartRow + template.locationNumberOffset, 3, 2, 1).clearContent();
      }
    }
  }
  
  // 写真URLを記録する配列
  const photoImageUrls = new Array(totalPhotos).fill('');
  
  for (let i = 0; i < totalPhotos; i++) {
    const photo = photos[i];
    if (photo.isBlank) continue;
    
    const dataStartRow = template.firstDataRow + i * template.blockDataSize;
    
    // 内容書き込み (C列に1行ずつ順番に)
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
      fieldsToWrite.push(`内容： ${photo.description}`);
    }
    if (displayFields.includes('testDetails') && photo.testFields) {
      const testKeys = Object.keys(photo.testFields);
      let hasTest = false;
      const testLines = ['【試験記録】'];
      for (let j = 0; j < testKeys.length; j++) {
        const key = testKeys[j];
        if (photo.testFields[key]) {
          testLines.push(`${getFieldLabel(key)}： ${photo.testFields[key]}`);
          hasTest = true;
        }
      }
      if (hasTest) {
        fieldsToWrite.push(...testLines);
      }
    }
    
    for (let j = 0; j < fieldsToWrite.length && j < template.contentRows; j++) {
      sheet.getRange(dataStartRow + j, 3).setValue(fieldsToWrite[j]);
    }
    
    // 撮影場所番号の書き込み
    if (template.hasLocationNumber && photo.locationNumber) {
      sheet.getRange(dataStartRow + template.locationNumberOffset, 3).setValue(photo.locationNumber);
    }
    
    // 写真の挿入
    if (photo.imageBase64) {
      try {
        const imageBlob = base64ToBlob(photo.imageBase64, `photo_${i + 1}`);
        const imageUrl = insertImageIntoCell(sheet, dataStartRow, template, imageBlob, photosFolder.getId());
        photoImageUrls[i] = imageUrl || '';
      } catch (err) {
        Logger.log('Image insertion error: ' + err.message);
      }
    } else if (photo.imageUrl) {
      // 既存のDrive画像URLを維持
      photoImageUrls[i] = photo.imageUrl;
    }
  }
  
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
  const directUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1600`;
  
  // 写真結合セルの行数 (例: A5:B24 の20行分)
  let rowSpan = 20;
  if (template && template.contentRows === 19) {
    rowSpan = 19;
  }
  const endRow = startRow + rowSpan - 1;
  
  // A列とB列の結合セル範囲(例: A5:B24)を再結合
  const photoRange = sheet.getRange(`A${startRow}:B${endRow}`);
  photoRange.merge();
  
  // スプレッドシート本体で100%確実に画像が即座描画される直リンクURL形式で=IMAGE数式をセット
  const anchorCell = sheet.getRange(`A${startRow}`);
  anchorCell.setFormula(`=IMAGE("${directUrl}", 1)`);
  
  return directUrl;
}

function getFieldLabel(key) {
  const labels = {
    'testPressure': '試験圧力',
    'holdTime': '保持時間',
    'startTime': '開始時間',
    'startPressure': '始圧',
    'waterLocation': '注水場所',
    'waterAmount': '注水量',
    'waterStatus': '採水状況',
  };
  return labels[key] || key;
}
