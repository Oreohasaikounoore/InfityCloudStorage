// レジストリ（ID保存用スプレッドシート）のID（既存のスプレッドシートのIDを指定）
var REGISTRY_ID = "あなたのスプレッドシートのIDを入力";
var PYTHONFILEID = "ここにドライブにアップロードしたPyファイルのIDを入力"

// 1セルあたりの最大文字数 MAX50000
var MAX_LENGTH = 49500;
// 進捗割合（パーセント）のキー
var PROGRESS_KEY = 'parsent';

/**
 * normalizeFileName: ファイル名中の URL エンコードされた文字（例："+" や "%20"）をスペースに戻し、
 * さらに保存時のファイル名に空白（半角・全角）が含まれないよう、すべて削除する。
 */
function normalizeFileName(name) {
  // URLエンコードされた文字（"+"など）をスペースに戻す
  var normalized = decodeURIComponent(name.replace(/\+/g, ' '));
  // すべての空白（半角・全角）を削除する
  normalized = normalized.replace(/\s/g, "");
  return normalized;
}



/**
 * doGet: ユーザーがアクセスしたときに GET パラメータ "filename" が指定されていれば、
 * そのファイルの base64 Data URL を取得しリダイレクトする。指定がなければ index.html を返す。
 */
function doGet(e) {
  if (e.parameter && e.parameter.filename) {
    var fileName = normalizeFileName(e.parameter.filename);
    var fileData = previewFile(fileName);
    if (fileData && fileData.data && fileData.type) {
      var base64Data = fileData.data.split(',')[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), fileData.type, fileName);    
      var redirectHtml = fileData.data;
      return ContentService.createTextOutput(redirectHtml)
                           .setMimeType(ContentService.MimeType.TEXT);
    } else {
      return ContentService.createTextOutput("File not found.");
    }
  } else {
    return HtmlService.createHtmlOutputFromFile('index');
  }
}



/**
 * getAllStorageSheets: レジストリシートに登録されている各スプレッドシートから、
 * シート名が "Data" で始まるすべてのシートを取得する。
 * @return {Array} オブジェクトの配列 [{ ss: Spreadsheet, sheet: Sheet }, ...]
 */
function getAllStorageSheets() {
  var regSS = SpreadsheetApp.openById(REGISTRY_ID);
  var regSheet = regSS.getSheets()[0];
  var lastRow = regSheet.getLastRow();
  if (lastRow < 1) return [];
  var data = regSheet.getRange(1, 1, lastRow, 1).getValues();
  var sheetsArr = [];
  for (var i = 0; i < data.length; i++) {
    var storageId = data[i][0];
    if (storageId) {
      try {
        var ss = SpreadsheetApp.openById(storageId);
        var allSheets = ss.getSheets();
        for (var j = 0; j < allSheets.length; j++) {
          if (allSheets[j].getName().indexOf("Data") === 0) {
            sheetsArr.push({ ss: ss, sheet: allSheets[j] });
          }
        }
      } catch (e) {
        Logger.log("Error opening spreadsheet: " + e.message);
      }
    }
  }
  return sheetsArr;
}



/**
 * adjustSheetFormatForSheet: 指定されたシートのフォーマットを調整する。
 * C列からZ列まで削除、行数を1000行に制限、A列とB列の幅を設定する。
 * @param {Sheet} sheet 対象のシート
 */
function adjustSheetFormatForSheet(sheet) {
  var currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols > 2) {
    sheet.deleteColumns(3, currentMaxCols - 2);
  }
  var currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows > 1000) {
    sheet.deleteRows(1001, currentMaxRows - 1000);
  }
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
}



/**
 * adjustSheetFormat: 新規作成したスプレッドシートの初期シートを "Data" に設定し、フォーマット調整する。
 * @param {Spreadsheet} ss 対象のスプレッドシート
 */
function adjustSheetFormat(ss) {
  var sheet = ss.getSheets()[0];
  sheet.setName("Data");
  adjustSheetFormatForSheet(sheet);
}



/**
 * getAvailableSheetInSpreadsheet: 指定のスプレッドシート内で、シート名が "Data" で始まるシートから
 * 空いている行があるシートを返す。もし既に「Data」シートが5枚以上なら null を返す。
 * 空きがなければ新しいシートを作成して返す。
 * @param {Spreadsheet} ss 対象のスプレッドシート
 * @return {Object|null} { sheet: Sheet, row: 利用可能な行番号 } または null
 */
function getAvailableSheetInSpreadsheet(ss) {
  var allSheets = ss.getSheets();
  var dataSheets = [];
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().indexOf("Data") === 0) {
      dataSheets.push(allSheets[i]);
    }
  }
  if (dataSheets.length >= 2) {
    return null;
  }
  for (var i = 0; i < dataSheets.length; i++) {
    var sheet = dataSheets[i];
    var maxRows = sheet.getMaxRows();
    var values = sheet.getRange(1, 1, maxRows, 1).getValues();
    for (var j = 0; j < values.length; j++) {
      if (!values[j][0]) {
        return { sheet: sheet, row: j + 1 };
      }
    }
  }
  var newSheetName = "Data" + (dataSheets.length + 1);
  var newSheet = ss.insertSheet(newSheetName);
  adjustSheetFormatForSheet(newSheet);
  return { sheet: newSheet, row: 1 };
}



/**
 * getStorageSheetForWriting: 登録済みスプレッドシートから、書き込み可能なシートを探す。
 * 書き込み可能なシートがなければ、新規スプレッドシートを作成して返す。
 * @return {Object} { ss: Spreadsheet, sheet: Sheet, row: 利用可能な行番号 }
 */
function getStorageSheetForWriting() {
  // レジストリが空の場合に備え、必要なストレージを作成しておく
  ensureRegistryNotEmpty();

  var regSS = SpreadsheetApp.openById(REGISTRY_ID);
  var regSheet = regSS.getSheets()[0];
  var regData = regSheet.getRange(1, 1, regSheet.getLastRow(), 1).getValues();
  for (var i = 0; i < regData.length; i++) {
    var storageId = regData[i][0];
    if (storageId) {
      var ss = SpreadsheetApp.openById(storageId);
      var available = getAvailableSheetInSpreadsheet(ss);
      if (available !== null) {
        return { ss: ss, sheet: available.sheet, row: available.row };
      }
    }
  }
  var newSS = SpreadsheetApp.create("InfCloud Data Storage " + new Date().toISOString());
  adjustSheetFormat(newSS);
  var newId = newSS.getId();
  regSheet.appendRow([newId]);
  SpreadsheetApp.flush();
  var newSheet = newSS.getSheetByName("Data");
  return { ss: newSS, sheet: newSheet, row: 1 };
}




/**
 * uploadData: ファイルアップロード処理  
 * data オブジェクト例: { name: "Test.png", type: "image/png", data: "base64文字列" }
 * ファイルを MAX_LENGTH ごとに分割し書き込み、進捗を ScriptProperties に更新する。
 */
function uploadData(data) {
  var fileName = normalizeFileName(data.name);
  if (fileName.includes(".py")) {
    return "成功";
  } else {
    try {
      try {
        deleteFile(fileName);
      } catch (e) {
        // 存在しなければ無視
      }
      var storage = getStorageSheetForWriting();
      var sheet = storage.sheet;
      var startRow = storage.row;
  
      var mimeType = data.type;
      var base64Data = data.data;
      var segments = [];
      for (var i = 0; i < base64Data.length; i += MAX_LENGTH) {
        segments.push(base64Data.substr(i, MAX_LENGTH));
      }
  
      PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, "0");
      var totalSegments = segments.length;
      for (var j = 0; j < totalSegments; j++) {
        if (startRow + j > sheet.getMaxRows()) {
          var available = getAvailableSheetInSpreadsheet(sheet.getParent());
          if (available === null) {
            break;
          }
          sheet = available.sheet;
          startRow = available.row;
        }
        if (j === 0) {
          sheet.getRange(startRow, 1).setValue('[' + fileName + ',' + mimeType + ']');
        } else {
          sheet.getRange(startRow + j, 1).setValue('[' + fileName + ',Next]');
        }
        sheet.getRange(startRow + j, 2).setValue(segments[j]);
        var progress = (j + 1) / totalSegments;
        PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, progress.toString());
      }
      
      // .py ファイル作成処理
      try {
        var folder = DriveApp.getFolderById(folderId);
        var existingFiles = folder.getFilesByName(fileName + ".py");
        while (existingFiles.hasNext()) {
          var existingFile = existingFiles.next();
          existingFile.setTrashed(true);
        }
        var savepy = fileName + ".py";
        var file = DriveApp.getFileById(PYTHONFILEID);
        var folder = DriveApp.getFolderById(folderId);
        file.makeCopy(savepy,folder);
      } catch (err) {
        Logger.log("Error creating .py file: " + err.toString());
      }
      return "完了";
    } catch (e) {
      return "アップロード失敗: " + e.toString();
    }
  }
}



/**
 * getfilename: 全登録済みスプレッドシートの「Data～」シートからファイル名一覧を取得する関数
 */
function getfilename(n) {
  try {
    // レジストリシートを取得
    var regSS = SpreadsheetApp.openById(REGISTRY_ID);
    var regSheet = regSS.getSheets()[0];
    var regData = regSheet.getRange(1, 1, regSheet.getLastRow(), 1).getValues();
    
    if (n === 0) {
      // 有効なスプレッドシートIDの数を返す
      var count = 0;
      for (var i = 0; i < regData.length; i++) {
        if (regData[i][0]) {
          count++;
        }
      }
      return count;
    } else {
      // nが1以上の場合、1-indexedでn番目のスプレッドシートを取得
      if (n > regData.length || !regData[n - 1][0]) {
        return "指定された番号のスプレッドシートは存在しません";
      }
      var storageId = regData[n - 1][0];
      var ss = SpreadsheetApp.openById(storageId);
      var sheets = ss.getSheets();
      var fileEntries = [];
      
      // 「Data」で始まるシート内のデータを取得
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().indexOf("Data") === 0) {
          var lastRow = sheets[i].getLastRow();
          if (lastRow === 0) continue;
          var data = sheets[i].getRange(1, 1, lastRow, 1).getValues();
          for (var j = 0; j < data.length; j++) {
            var cell = data[j][0];
            if (!cell) continue;
            var match = cell.match(/^\[(.*?),(.*?)\]$/);
            // 「Next」以外のエントリかつ重複しないものだけを追加
            if (match && match[2] !== "Next") {
              var exists = false;
              for (var k = 0; k < fileEntries.length; k++) {
                if (fileEntries[k].name === match[1]) {
                  exists = true;
                  break;
                }
              }
              if (!exists) {
                fileEntries.push({ name: match[1], storageId: storageId });
              }
            }
          }
        }
      }
      return fileEntries;
    }
  } catch (e) {
    return "取得できませんでした: " + e.toString();
  }
}

/**
 * deleteFile: 指定したファイル名のファイルを全登録済みスプレッドシートから削除し、
 * ドライブ上の同名の .py ファイルも削除する関数  
 * ※削除後、各シートの行数が1000行未満なら不足分の行を追加して常に1000行に保つ。
 */
function deleteFile(fileName) {
  fileName = normalizeFileName(fileName);
  try {
    var sheetsArr = getAllStorageSheets();
    var foundAny = false;
    
    sheetsArr.forEach(function(obj) {
      var sheet = obj.sheet;
      var maxRows = sheet.getMaxRows();
      var range = sheet.getRange(1, 1, maxRows, 2);
      var values = range.getValues();
      var rowsToDelete = [];
      
      for (var i = 0; i < values.length; i++) {
        var cell = values[i][0];
        if (cell) {
          var match = cell.match(/^\[(.*?),(.*?)\]$/);
          if (match && match[1] === fileName) {
            rowsToDelete.push(i + 1);
          }
        }
      }
      
      if (rowsToDelete.length > 0) {
        foundAny = true;
        var lastRow = sheet.getLastRow();
        if (rowsToDelete.length === lastRow) {
          rowsToDelete.forEach(function(row) {
            sheet.getRange(row, 1, 1, 2).clearContent();
          });
        } else {
          rowsToDelete.sort(function(a, b) { return b - a; });
          rowsToDelete.forEach(function(row) {
            sheet.deleteRow(row);
          });
        }
      }
      
      // 削除後、行数が1000未満なら不足分を追加
      var currentRows = sheet.getMaxRows();
      if (currentRows < 1000) {
        sheet.insertRowsAfter(currentRows, 1000 - currentRows);
      }
    });
    
    try {
      var folder = DriveApp.getFolderById(folderId);
      var pyFiles = folder.getFilesByName(fileName + ".py");
      while (pyFiles.hasNext()) {
        var pyFile = pyFiles.next();
        pyFile.setTrashed(true);
      }
    } catch (e) {
      Logger.log("Drive上の .py ファイル削除に失敗: " + e.toString());
    }
    
    return foundAny ? "削除完了" : "ファイルが存在しません";
  } catch (e) {
    return "削除失敗: " + e.toString();
  }
}



/**
 * getCapacity: 全登録済みスプレッドシートのデータ容量（base64データから概算）を合算して返す関数
 */
function getCapacityParts(n) {
  try {
    // レジストリシートを取得
    var regSS = SpreadsheetApp.openById(REGISTRY_ID);
    var regSheet = regSS.getSheets()[0];
    var regData = regSheet.getRange(1, 1, regSheet.getLastRow(), 1).getValues();
    
    if (n === 0) {
      // 有効なストレージ（スプレッドシート）の数を返す
      var count = 0;
      for (var i = 0; i < regData.length; i++) {
        if (regData[i][0]) count++;
      }
      return count;
    } else {
      // n が 1 以上の場合、1-indexed で n 番目のストレージを取得
      if (n > regData.length || !regData[n - 1][0]) {
        return "指定された番号のスプレッドシートは存在しません";
      }
      var storageId = regData[n - 1][0];
      var ss = SpreadsheetApp.openById(storageId);
      var sheets = ss.getSheets();
      var totalCapacity = 0;
      
      // 各シートのうち、名前が "Data" で始まるシートを対象に容量を合算
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().indexOf("Data") === 0) {
          var maxRows = sheets[i].getMaxRows();
          // データは2列目に格納されている前提
          var data = sheets[i].getRange(1, 2, maxRows, 1).getValues();
          for (var j = 0; j < data.length; j++) {
            var cellData = data[j][0];
            if (cellData) {
              // base64 の文字数から概算バイト数を計算（4文字で3バイト）
              totalCapacity += Math.floor(cellData.length * 3 / 4);
            }
          }
        }
      }
      return totalCapacity;
    }
  } catch (e) {
    return "取得失敗: " + e.toString();
  }
}


/**
 * previewFile: 指定したファイル名のファイルを全登録済みスプレッドシートから検索し、
 * base64データを再構築して data URI として返す関数
 */
function previewFile(fileName, storageId) {
  try {
    var sheetsArr = [];
    if (storageId) {
      var ss = SpreadsheetApp.openById(storageId);
      var allSheets = ss.getSheets();
      for (var j = 0; j < allSheets.length; j++) {
        if (allSheets[j].getName().indexOf("Data") === 0) {
          sheetsArr.push({ sheet: allSheets[j] });
        }
      }
    } else {
      sheetsArr = getAllStorageSheets();
    }
    var base64Data = "";
    var mimeType = "";
    var found = false;
    for (var k = 0; k < sheetsArr.length; k++) {
      var sheet = sheetsArr[k].sheet;
      var maxRows = sheet.getMaxRows();
      var data = sheet.getRange(1, 1, maxRows, 2).getValues();
      for (var i = 0; i < data.length; i++) {
        var cell = data[i][0];
        if (cell) {
          var match = cell.match(/^\[(.*?),(.*?)\]$/);
          if (match && match[1] === fileName) {
            if (!found) {
              mimeType = match[2];
              base64Data += data[i][1];
              found = true;
            } else if (match[2] === "Next") {
              base64Data += data[i][1];
            }
          }
        }
      }
      if (found) break;
    }
    if (!base64Data) {
      return { type: '', data: '' };
    }
    var dataUrl = "data:" + mimeType + ";base64," + base64Data;
    return { type: mimeType, data: dataUrl };
  } catch (e) {
    return { type: '', data: '' };
  }
}


/**
 * getProgress: 現在のアップロード進捗（割合）を ScriptProperties から取得して返す関数
 */
function getProgress() {
  return PropertiesService.getScriptProperties().getProperty(PROGRESS_KEY) || "0";
}



/**
 * saveDriveFilesToInfCloud: Google Drive の特定フォルダ内のファイルを InfCloud に保存し削除する関数  
 * ※処理は1ファイルのみ実行し、かつ、対象ファイルが .py ファイルの場合はフォルダ内に
 *    .py 以外のファイルが存在する限りスキップして、.py 以外のファイルを優先的に処理します。
 */
function saveDriveFilesToInfCloud() {
  var folder = DriveApp.getFolderById(folderId);
  
  // すべてのファイルを配列に格納
  var filesIterator = folder.getFiles();
  var filesArray = [];
  while (filesIterator.hasNext()) {
    filesArray.push(filesIterator.next());
  }
  
  // フォルダ内に、.py 以外のファイルがあるかチェック
  var hasNonPy = filesArray.some(function(file) {
    return !normalizeFileName(file.getName()).toLowerCase().endsWith(".py");
  });
  
  // 1ファイルのみ処理（優先は .py 以外のファイル）
  for (var i = 0; i < filesArray.length; i++) {
    var file = filesArray[i];
    var fileName = normalizeFileName(file.getName());
    
    // .py ファイルで、かつ非 .py ファイルが存在するならスキップする
    if (fileName.toLowerCase().endsWith(".py") && hasNonPy) {
      continue;
    }
    
    // ここで対象ファイル（非 .py または、非 .py がなければ .py）を1件処理する
    var mimeType = file.getMimeType();
    var blob = file.getBlob();
    var base64Data = Utilities.base64Encode(blob.getBytes());
  
    var data = {
      name: fileName,
      type: mimeType,
      data: base64Data
    };
  
    var result = uploadData(data);
  
    if (result === "完了") {
      file.setTrashed(true);
      try {
        var fileData = DriveApp.getFilesByName(fileName + ".py");
        if (fileData.hasNext()) {
          fileData.next().setTrashed(true);
        }
      } catch (e) {
        console.log("pyファイルは存在しません");
      }
      var savepy = fileName + ".py";
      var file = DriveApp.getFileById(PYTHONFILEID);
      var folder = DriveApp.getFolderById(folderId);
      file.makeCopy(savepy,folder);
      Logger.log("Successfully uploaded and moved to trash: " + fileName);
    } else {
      Logger.log("Failed to save: " + fileName);
    }
    // 1件処理したらループを抜ける
    break;
  }
}
function deleteAllData() {
  try {
    var sheetsArr = getAllStorageSheets();
    sheetsArr.forEach(function(obj) {
      var sheet = obj.sheet;
      var maxRows = sheet.getMaxRows();
      // 全データをクリア（1列目と2列目）
      sheet.getRange(1, 1, maxRows, 2).clearContent();
      // 1000行未満の場合、追加する
      var currentRows = sheet.getMaxRows();
      if (currentRows < 1000) {
        sheet.insertRowsAfter(currentRows, 1000 - currentRows);
      }
    });
    
    // ※必要に応じて、同じフォルダ内の .py ファイルも削除
    try {
      var folder = DriveApp.getFolderById(folderId);
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName().toLowerCase().endsWith(".py")) {
          file.setTrashed(true);
        }
      }
    } catch (e) {
      Logger.log("pyファイル削除エラー: " + e.toString());
    }
    
    return "全データ削除完了";
  } catch (e) {
    return "全データ削除失敗: " + e.toString();
  }
}
/**
 * ensureRegistryNotEmpty:
 * レジストリシートが空の場合、新規ストレージ用スプレッドシートを作成してIDを追加する関数
 */
function ensureRegistryNotEmpty() {
  var regSS = SpreadsheetApp.openById(REGISTRY_ID);
  var regSheet = regSS.getSheets()[0];
  if (regSheet.getLastRow() < 1) {
    var newSS = SpreadsheetApp.create("InfCloud Data Storage " + new Date().toISOString());
    adjustSheetFormat(newSS);
    var newId = newSS.getId();
    regSheet.appendRow([newId]);
    SpreadsheetApp.flush();
  }
}
