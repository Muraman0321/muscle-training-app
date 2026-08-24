/**
 * 筋トレ進捗管理アプリ - GAS Web App backend
 *
 * デプロイ手順は apps-script/SETUP.md を参照。
 *
 * シート構成:
 *   種目マスタ     : ID, 種目名, 部位, 作成日
 *   トレーニングログ : 日時, 種目名, 重量kg, 回数, セット, メモ
 *   体組成ログ     : 日時, 体重kg, 体脂肪率, メモ
 *   長期メニュー   : 週, 曜日, 種目名, 目標セット, 目標回数, 目標重量, メモ
 */

var SHEET_EXERCISES = '種目マスタ';
var SHEET_TRAINING = 'トレーニングログ';
var SHEET_BODY = '体組成ログ';
var SHEET_MENU = '長期メニュー';

var HEADERS = {};
HEADERS[SHEET_EXERCISES] = ['ID', '種目名', '部位', '作成日'];
HEADERS[SHEET_TRAINING] = ['日時', '種目名', '重量kg', '回数', 'セット', 'メモ'];
HEADERS[SHEET_BODY] = ['日時', '体重kg', '体脂肪率', 'メモ'];
HEADERS[SHEET_MENU] = ['週', '曜日', '種目名', '目標セット', '目標回数', '目標重量', 'メモ'];

/**
 * Apps Scriptエディタから手動で一度実行する。
 * 4シートを(なければ)作成し、ヘッダー行を設定する。
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    var headers = HEADERS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  });
  // デフォルトで作られる「シート1」は使わないので、他が揃っていれば削除
  var defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 4) {
    ss.deleteSheet(defaultSheet);
  }
  Logger.log('setup done. Script Properties に SECRET を設定するのを忘れずに。');
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function checkToken_(token) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  return secret && token === secret;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = values.slice(1);
  return rows
    .filter(function (row) {
      return row.some(function (cell) {
        return cell !== '' && cell !== null;
      });
    })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) {
        var v = row[i];
        if (v instanceof Date) {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        }
        obj[h] = v;
      });
      return obj;
    });
}

function ensureExercise_(name, bodyPart) {
  if (!name) return;
  var sheet = getSheet_(SHEET_EXERCISES);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === name) return;
  }
  var nextId = data.length; // header込みなので data.length = 次のID(1始まり)
  sheet.appendRow([nextId, name, bodyPart || '', new Date()]);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'ping';

  if (action === 'ping') {
    return jsonOut_({ ok: true, message: 'muscle-training-app GAS is alive' });
  }

  if (!checkToken_(params.token)) {
    return jsonOut_({ ok: false, error: 'invalid token' });
  }

  try {
    if (action === 'listExercises') {
      return jsonOut_({ ok: true, data: sheetToObjects_(getSheet_(SHEET_EXERCISES)) });
    }
    if (action === 'listSets') {
      return jsonOut_({ ok: true, data: sheetToObjects_(getSheet_(SHEET_TRAINING)) });
    }
    if (action === 'listBody') {
      return jsonOut_({ ok: true, data: sheetToObjects_(getSheet_(SHEET_BODY)) });
    }
    if (action === 'listMenu') {
      return jsonOut_({ ok: true, data: sheetToObjects_(getSheet_(SHEET_MENU)) });
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid JSON body' });
  }

  if (!checkToken_(body.token)) {
    return jsonOut_({ ok: false, error: 'invalid token' });
  }

  var action = body.action;
  var payload = body.payload;

  try {
    if (action === 'addSets') {
      return handleAddSets_(payload);
    }
    if (action === 'addBody') {
      return handleAddBody_(payload);
    }
    if (action === 'addExercise') {
      return handleAddExercise_(payload);
    }
    if (action === 'deleteExercise') {
      return handleDeleteExercise_(payload);
    }
    if (action === 'importMenu') {
      return handleImportMenu_(payload);
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// payload: [{ date, exercise, weight, reps, setNumber, memo }, ...]
function handleAddSets_(payload) {
  if (!Array.isArray(payload) || payload.length === 0) {
    return jsonOut_({ ok: false, error: 'payload must be a non-empty array' });
  }
  var sheet = getSheet_(SHEET_TRAINING);
  var rows = payload.map(function (s) {
    ensureExercise_(s.exercise);
    return [
      s.date ? new Date(s.date) : new Date(),
      s.exercise || '',
      Number(s.weight) || 0,
      Number(s.reps) || 0,
      Number(s.setNumber) || 1,
      s.memo || '',
    ];
  });
  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
  return jsonOut_({ ok: true, added: rows.length });
}

// payload: { date, weight, bodyFat, memo }
function handleAddBody_(payload) {
  var sheet = getSheet_(SHEET_BODY);
  sheet.appendRow([
    payload.date ? new Date(payload.date) : new Date(),
    Number(payload.weight) || '',
    payload.bodyFat !== undefined && payload.bodyFat !== '' ? Number(payload.bodyFat) : '',
    payload.memo || '',
  ]);
  return jsonOut_({ ok: true });
}

// payload: { name, bodyPart }
function handleAddExercise_(payload) {
  ensureExercise_(payload.name, payload.bodyPart);
  return jsonOut_({ ok: true });
}

// payload: { name }
function handleDeleteExercise_(payload) {
  var sheet = getSheet_(SHEET_EXERCISES);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === payload.name) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return jsonOut_({ ok: true });
}

// payload: [{ week, day, exercise, sets, reps, weight, memo }, ...]
// 長期メニューは毎回全置き換え(既存行を消してから書き込む)
function handleImportMenu_(payload) {
  if (!Array.isArray(payload)) {
    return jsonOut_({ ok: false, error: 'payload must be an array' });
  }
  var sheet = getSheet_(SHEET_MENU);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  if (payload.length === 0) {
    return jsonOut_({ ok: true, imported: 0 });
  }
  var rows = payload.map(function (m) {
    return [
      Number(m.week) || '',
      m.day || '',
      m.exercise || '',
      Number(m.sets) || '',
      m.reps || '',
      m.weight || '',
      m.memo || '',
    ];
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  return jsonOut_({ ok: true, imported: rows.length });
}
