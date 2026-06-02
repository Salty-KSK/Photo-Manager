/**
 * 工事写真台帳 - Google Apps Script バックエンド
 * 
 * このスクリプトをGASエディタにコピー＆ペーストし、
 * Webアプリとしてデプロイしてください。
 * 
 * デプロイ手順:
 * 1. https://script.google.com/ で新しいプロジェクトを作成
 * 2. このコードを貼り付け
 * 3. TEMPLATE_SPREADSHEET_ID を設定（ひな型スプレッドシートのID）
 * 4. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
 * 5. 「次のユーザーとして実行」を「自分」に設定
 * 6. 「アクセスできるユーザー」を「全員」に設定
 * 7. デプロイ後に表示されるURLをReactアプリのGAS_URLに設定
 */

// ひな形スプレッドシートのID
const TEMPLATE_SPREADSHEET_ID = '1DfZzuhZN55BTCkg9li2ubmjqCwsS59AHEI5aIhFPEtw';

// 写真を保存するGoogle DriveフォルダのID（空の場合はマイドライブ直下）
const PHOTO_FOLDER_ID = '';

// 1ブロックあたりの行数
const BLOCK_SIZE = 21; // ブロックの実サイズは21行であることが判明！

// ▼ いただいた完璧なセル配置（Top-Left起点）に合わせて修正 ▼
const CELL_OFFSETS = {
  projectNameLine1: 0,  // 行0 (1行目基準で E1, E22...)
  image: 1,             // 行1 (1行目基準で B2, B23...)
  date: 3,              // 行3 (1行目基準で E4, E25...) 
  location: 5,          // 行5 (1行目基準で E6, E27...) 
  category: 7,          // 行7 (1行目基準で E8, E29...) 
  description: 9,       // 行9 (1行目基準で E10, E31...) 
};

function formatJapaneseDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

/**
 * GETリクエスト処理（インポート用）
 */
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'import') {
    const spreadsheetId = e.parameter.spreadsheetId;
    return handleImport(spreadsheetId);
  } else if (action === 'list') {
    return handleList();
  }
  
  return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POSTリクエスト処理（エクスポート用）
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.action === 'export') {
      return handleExport(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * エクスポート処理：テンプレートをコピーしてデータを書き込む
 */
function handleExport(data) {
  const { projectNameLine1, projectNameLine2, photos } = data;
  
  // フォルダ名の生成
  const todayStr = new Date().toLocaleDateString('ja-JP').replace(/\//g, '');
  const folderName = `工事写真台帳_${projectNameLine1 || ''}${projectNameLine2 || ''}_${todayStr}`;
  
  // 親フォルダと写真フォルダの作成
  let baseFolder = PHOTO_FOLDER_ID ? DriveApp.getFolderById(PHOTO_FOLDER_ID) : DriveApp.getRootFolder();
  const mainFolder = baseFolder.createFolder(folderName);
  const photosFolder = mainFolder.createFolder('写真');
  
  const templateFile = DriveApp.getFileById(TEMPLATE_SPREADSHEET_ID);
  const newFile = templateFile.makeCopy(folderName, mainFolder);
  const ss = SpreadsheetApp.openById(newFile.getId());
  const sheet = ss.getSheets()[0];
  
  // 2. 3枚を超える場合、必要なページ分のブロックを追加
  const totalPhotos = photos.length;
  const blocksPerPage = 3;
  const totalBlocksNeeded = totalPhotos;
  const existingBlocks = blocksPerPage; // テンプレートには3ブロック分ある
  
  if (totalBlocksNeeded > existingBlocks) {
    // ブロック1全体の範囲を取得 (安全のためA〜Lの12列までに制限し、余分なマージの交差を防ぐ)
    const block1Range = sheet.getRange(1, 1, BLOCK_SIZE, 12);
    
    for (let i = existingBlocks; i < totalBlocksNeeded; i++) {
      const startRow = i * BLOCK_SIZE + 1;
      
      // 行を挿入
      sheet.insertRowsAfter(startRow - 1, BLOCK_SIZE);
      
      // ブロック1の書式ごとコピー
      const targetRange = sheet.getRange(startRow, 1);
      block1Range.copyTo(targetRange);
      
      // コピー元のデータ(文字・画像)だけをクリア（結合セルのエラーを防ぐためピンポイントでクリア）
      // ▼ コピー元データのクリア ▼
      sheet.getRange(`B${startRow + CELL_OFFSETS.image}`).clearContent(); 
      sheet.getRange(`E${startRow + CELL_OFFSETS.projectNameLine1}`).clearContent();
      sheet.getRange(`E${startRow + CELL_OFFSETS.date}`).clearContent();
      sheet.getRange(`E${startRow + CELL_OFFSETS.location}`).clearContent();
      sheet.getRange(`E${startRow + CELL_OFFSETS.category}`).clearContent();
      sheet.getRange(`E${startRow + CELL_OFFSETS.description}`).clearContent();
    }
  }
  
  // 3. 各写真のデータを書き込み
  for (let i = 0; i < totalPhotos; i++) {
    const photo = photos[i];
    const blockStartRow = i * BLOCK_SIZE + 1;
    
    // 工事名称 (B1などのセルに改行して2項目を入れる)
    const fullName = [projectNameLine1, projectNameLine2].filter(Boolean).join('\n');
    sheet.getRange(`E${blockStartRow + CELL_OFFSETS.projectNameLine1}`).setValue(fullName);
    
    if (photo.isBlank) continue; // 空白ブロックはスキップ
    
    // 同じセルにラベルと値をセットで書き込む
    const formattedDate = formatJapaneseDate(photo.date);
    sheet.getRange(`E${blockStartRow + CELL_OFFSETS.date}`).setValue(`日付： ${formattedDate}`);
    sheet.getRange(`E${blockStartRow + CELL_OFFSETS.location}`).setValue(`場所： ${photo.location || ''}`);
    sheet.getRange(`E${blockStartRow + CELL_OFFSETS.category}`).setValue(`種別： ${photo.category || ''}`);
    
    // 内容 ＋ 試験詳細フィールドの統合
    let finalDescription = photo.description || '';
    if (photo.testFields) {
      const fieldKeys = Object.keys(photo.testFields);
      let testNotes = [];
      for (let j = 0; j < fieldKeys.length; j++) {
        const key = fieldKeys[j];
        const value = photo.testFields[key];
        if (value) {
          testNotes.push(`${getFieldLabel(key)}: ${value}`);
        }
      }
      if (testNotes.length > 0) {
        finalDescription += (finalDescription ? '\n\n' : '') + '【試験記録】\n' + testNotes.join('\n');
      }
    }
    sheet.getRange(`E${blockStartRow + CELL_OFFSETS.description}`).setValue(`内容： ${finalDescription}`);
    
    // 写真を挿入（Base64データがある場合）
    if (photo.imageBase64) {
      try {
        const imageBlob = base64ToBlob(photo.imageBase64, `photo_${i + 1}`);
        // 画像はセルBの offset位置に保存
        insertImageIntoCell(sheet, `B${blockStartRow + CELL_OFFSETS.image}`, imageBlob, photosFolder.getId());
      } catch (err) {
        Logger.log('Image insertion error: ' + err.message);
      }
    }
  }
  
  // 4. _metadataシートにJSONデータを保存（再編集用）
  let metaSheet = ss.getSheetByName('_metadata');
  if (!metaSheet) {
    metaSheet = ss.insertSheet('_metadata');
  }
  
  const metaData = {
    projectNameLine1,
    projectNameLine2,
    photos: photos.map(p => ({
      date: p.date,
      location: p.location,
      category: p.category,
      testType: p.testType,
      description: p.description,
      testFields: p.testFields,
      isBlank: p.isBlank,
    })),
    exportDate: new Date().toISOString(),
  };
  
  metaSheet.getRange('A1').setValue(JSON.stringify(metaData));
  metaSheet.hideSheet();
  
  // 5. URLを返却
  const url = ss.getUrl();
  
  return ContentService.createTextOutput(JSON.stringify({ 
    success: true, 
    url: url,
    spreadsheetId: newFile.getId()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * インポート処理：_metadataシートからデータを読み取る
 */
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

/**
 * 最近作成された台帳リストを取得する
 */
function handleList() {
  try {
    const query = "title contains '工事写真台帳_' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false";
    const files = DriveApp.searchFiles(query);
    const fileList = [];
    
    // 最近作成されたもの上位20件程度を取得
    let count = 0;
    while (files.hasNext() && count < 20) {
      const file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        date: file.getDateCreated().getTime() // ソート用にタイムスタンプ
      });
      count++;
    }
    
    // 日付の降順（新しい順）にソート
    fileList.sort((a, b) => b.date - a.date);
    
    // フォーマット変換
    const formattedList = fileList.map(item => {
      const d = new Date(item.date);
      // YYYY/MM/DD HH:mm 形式
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      return {
        id: item.id,
        name: item.name,
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

/**
 * Base64データをBlobに変換
 */
function base64ToBlob(base64Data, fileName) {
  const parts = base64Data.split(',');
  const contentType = parts[0].match(/:(.*?);/)[1];
  const raw = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(raw, contentType, fileName);
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

/**
 * セルに画像を挿入
 */
function insertImageIntoCell(sheet, cellRef, imageBlob, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(imageBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  const imageUrl = `https://drive.google.com/uc?id=${file.getId()}`;
  const cell = sheet.getRange(cellRef);
  
  try {
    // IMAGE関数を使用せず、直接セルに画像を埋め込むことでアクセス許可の警告を回避する
    const image = SpreadsheetApp.newCellImage().setSourceUrl(imageUrl).build();
    cell.setValue(image);
  } catch (e) {
    // 予期せぬエラー時のフォールバック
    cell.setFormula(`=IMAGE("${imageUrl}", 1)`);
  }
}

/**
 * テストフィールドのキーからラベルを取得
 */
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
