/**
 * 百問計算 記録・グラフ化 Webアプリ
 * バックエンド: Google Apps Script
 * 
 * スプレッドシート構成:
 * - 「ユーザー」シート: A:出席番号, B:氏名, C:ログインID, D:パスワード, E:権限
 * - 「記録」シート: A:タイムスタンプ, B:出席番号, C:種目, D:プリント番号, E:正答数, F:タイム(秒)
 */

// =============================================
// 定数
// =============================================
const SHEET_USERS = 'ユーザー';
const SHEET_RECORDS = '記録';
const COL_USER_SEAT = 1;     // A: 出席番号
const COL_USER_NAME = 2;     // B: 氏名
const COL_USER_ID = 3;       // C: ログインID
const COL_USER_PASS = 4;     // D: パスワード
const COL_USER_ROLE = 5;     // E: 権限
const COL_REC_TIMESTAMP = 1; // A: タイムスタンプ
const COL_REC_SEAT = 2;      // B: 出席番号
const COL_REC_SUBJECT = 3;   // C: 種目
const COL_REC_PRINT = 4;     // D: プリント番号
const COL_REC_SCORE = 5;     // E: 正答数
const COL_REC_TIME = 6;      // F: タイム(秒)

// =============================================
// Webアプリ エントリーポイント
// =============================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('百問計算 記録アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================
// スプレッドシート ユーティリティ
// =============================================
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getUserSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    // ヘッダー行を追加
    sheet.appendRow(['出席番号', '氏名', 'ログインID', 'パスワード', '権限']);
  }
  return sheet;
}

function getRecordSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RECORDS);
    // ヘッダー行を追加
    sheet.appendRow(['タイムスタンプ', '出席番号', '種目', 'プリント番号', '正答数', 'タイム(秒)']);
  }
  return sheet;
}

function getUserData() {
  const sheet = getUserSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 5).getValues();
}

function getRecordData() {
  const sheet = getRecordSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 6).getValues();
}

// =============================================
// 認証機能
// =============================================

/**
 * ログイン処理
 * @param {string} loginId - ログインID
 * @param {string} password - パスワード
 * @returns {Object} - {success, user: {seatNo, name, loginId, role}, message}
 */
function login(loginId, password) {
  try {
    const users = getUserData();
    for (let i = 0; i < users.length; i++) {
      const row = users[i];
      const storedId = String(row[COL_USER_ID - 1]).trim();
      const storedPass = row[COL_USER_PASS - 1]; // 文字列変換せずに取得
      
      if (storedId === String(loginId).trim()) {
        // 初回登録（パスワード未設定、または空文字、null、空白のみの場合）
        if (storedPass === '' || storedPass === null || storedPass === undefined || String(storedPass).trim() === '') {
          return {
            success: false,
            needsPasswordSetup: true,
            user: {
              seatNo: row[COL_USER_SEAT - 1],
              name: String(row[COL_USER_NAME - 1]),
              loginId: storedId,
              role: String(row[COL_USER_ROLE - 1])
            },
            message: '初回ログインです。パスワードを設定してください。'
          };
        }
        if (storedPass === String(password).trim()) {
          return {
            success: true,
            user: {
              seatNo: row[COL_USER_SEAT - 1],
              name: String(row[COL_USER_NAME - 1]),
              loginId: storedId,
              role: String(row[COL_USER_ROLE - 1])
            },
            message: 'ログイン成功'
          };
        } else {
          return { success: false, message: 'パスワードが違います。' };
        }
      }
    }
    return { success: false, message: 'ログインIDが見つかりません。' };
  } catch (e) {
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  }
}

/**
 * 初回パスワード登録
 * @param {string} loginId - ログインID
 * @param {string} newPassword - 新しいパスワード
 * @returns {Object} - {success, message}
 */
function registerPassword(loginId, newPassword) {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    for (let i = 0; i < users.length; i++) {
      const storedId = String(users[i][COL_USER_ID - 1]).trim();
      if (storedId === String(loginId).trim()) {
        const rowNum = i + 2; // ヘッダー行分 +1、0-indexed分 +1
        sheet.getRange(rowNum, COL_USER_PASS).setValue(String(newPassword).trim());
        return { success: true, message: 'パスワードを設定しました。' };
      }
    }
    return { success: false, message: 'ユーザーが見つかりません。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// =============================================
// 記録機能 (児童用)
// =============================================

/**
 * 記録を保存する
 * @param {Object} data - {seatNo, subject, printNo, wrongCount, minutes, seconds}
 * @returns {Object} - {success, message}
 */
function saveRecord(data) {
  try {
    const sheet = getRecordSheet();
    const score = 100 - parseInt(data.wrongCount, 10);
    const totalSeconds = parseInt(data.minutes, 10) * 60 + parseInt(data.seconds, 10);
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    
    sheet.appendRow([
      timestamp,
      data.seatNo,
      data.subject,
      data.printNo,
      score,
      totalSeconds
    ]);
    
    return { success: true, message: '記録を保存しました！ 正答数: ' + score + '問', score: score };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 自分の記録を取得する
 * @param {number|string} seatNo - 出席番号
 * @returns {Object} - {success, records: [{date, subject, printNo, score, time}]}
 */
function getMyRecords(seatNo) {
  try {
    const records = getRecordData();
    const myRecords = records
      .filter(row => String(row[COL_REC_SEAT - 1]) === String(seatNo))
      .map(row => ({
        date: formatDate(row[COL_REC_TIMESTAMP - 1]),
        subject: row[COL_REC_SUBJECT - 1],
        printNo: row[COL_REC_PRINT - 1],
        score: row[COL_REC_SCORE - 1],
        time: row[COL_REC_TIME - 1]
      }))
      .sort((a, b) => new Date(a.date) - new Date(b.date));
    
    return { success: true, records: myRecords };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message, records: [] };
  }
}

/**
 * クラス全体の統計（日付・種目ごとの平均・最速タイム）を取得する
 * @returns {Object} - {success, stats: {[date_subject]: {avg, best}}}
 */
function getClassStats() {
  try {
    const records = getRecordData();
    const statsMap = {};
    
    records.forEach(row => {
      const date = formatDate(row[COL_REC_TIMESTAMP - 1]);
      const subject = row[COL_REC_SUBJECT - 1];
      const time = Number(row[COL_REC_TIME - 1]);
      if (!date || !subject || isNaN(time) || time <= 0) return;
      
      const key = date + '__' + subject;
      if (!statsMap[key]) {
        statsMap[key] = { date, subject, times: [] };
      }
      statsMap[key].times.push(time);
    });
    
    const stats = {};
    Object.keys(statsMap).forEach(key => {
      const { date, subject, times } = statsMap[key];
      const avg = Math.round(times.reduce((a, b) => a + b, 0) / times.length);
      const best = Math.min(...times);
      stats[key] = { date, subject, avg, best };
    });
    
    return { success: true, stats };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message, stats: {} };
  }
}

// =============================================
// 名簿管理機能 (教師用)
// =============================================

/**
 * 児童名簿を取得する
 * @returns {Object} - {success, students: [{seatNo, name, loginId, role}]}
 */
function getStudentList() {
  try {
    const users = getUserData();
    const students = users
      .filter(row => String(row[COL_USER_ROLE - 1]) === '児童')
      .map(row => ({
        seatNo: row[COL_USER_SEAT - 1],
        name: String(row[COL_USER_NAME - 1]),
        loginId: String(row[COL_USER_ID - 1]),
        role: String(row[COL_USER_ROLE - 1])
      }))
      .sort((a, b) => Number(a.seatNo) - Number(b.seatNo));
    
    return { success: true, students };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message, students: [] };
  }
}

/**
 * 児童名簿を一括保存する（Excelアップロード用）
 * @param {Array} students - [{seatNo, name}]
 * @returns {Object} - {success, message, count}
 */
function saveStudentList(students) {
  try {
    const sheet = getUserSheet();
    let addedCount = 0;
    
    students.forEach(student => {
      const seatNo = student.seatNo;
      const name = student.name;
      if (!seatNo || !name) return;
      
      // 既存の出席番号を検索
      const users = getUserData();
      const existingIndex = users.findIndex(row => String(row[COL_USER_SEAT - 1]) === String(seatNo));
      
      if (existingIndex >= 0) {
        // 上書き更新
        const rowNum = existingIndex + 2;
        sheet.getRange(rowNum, COL_USER_SEAT).setValue(seatNo);
        sheet.getRange(rowNum, COL_USER_NAME).setValue(name);
        // ログインIDが未設定なら自動生成
        if (!sheet.getRange(rowNum, COL_USER_ID).getValue()) {
          sheet.getRange(rowNum, COL_USER_ID).setValue('student' + String(seatNo).padStart(2, '0'));
          sheet.getRange(rowNum, COL_USER_ROLE).setValue('児童');
        }
      } else {
        // 新規追加
        sheet.appendRow([
          seatNo,
          name,
          'student' + String(seatNo).padStart(2, '0'),
          '',
          '児童'
        ]);
        addedCount++;
      }
    });
    
    return { success: true, message: addedCount + '名を追加しました。', count: addedCount };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 児童を個別登録・更新する
 * @param {Object} student - {seatNo, name, loginId}
 * @returns {Object} - {success, message}
 */
function addOrUpdateStudent(student) {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    const seatNo = String(student.seatNo).trim();
    const name = String(student.name).trim();
    const loginId = student.loginId ? String(student.loginId).trim() : 'student' + seatNo.padStart(2, '0');

    if (!seatNo || !name) {
      return { success: false, message: '出席番号と氏名は必須です。' };
    }

    const existingIndex = users.findIndex(row => String(row[COL_USER_SEAT - 1]) === seatNo);
    
    if (existingIndex >= 0) {
      const rowNum = existingIndex + 2;
      sheet.getRange(rowNum, COL_USER_SEAT).setValue(seatNo);
      sheet.getRange(rowNum, COL_USER_NAME).setValue(name);
      sheet.getRange(rowNum, COL_USER_ID).setValue(loginId);
      sheet.getRange(rowNum, COL_USER_ROLE).setValue('児童');
      return { success: true, message: '更新しました。' };
    } else {
      sheet.appendRow([seatNo, name, loginId, '', '児童']);
      return { success: true, message: '登録しました。' };
    }
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 児童を個別削除する
 * @param {string|number} seatNo - 出席番号
 * @returns {Object} - {success, message}
 */
function deleteStudent(seatNo) {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    const targetIndex = users.findIndex(row => String(row[COL_USER_SEAT - 1]) === String(seatNo));
    
    if (targetIndex < 0) {
      return { success: false, message: '対象の児童が見つかりません。' };
    }
    
    const rowNum = targetIndex + 2;
    sheet.deleteRow(rowNum);
    return { success: true, message: '削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 全児童を一括削除する
 * @returns {Object} - {success, message}
 */
function deleteAllStudents() {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    // 後ろから削除（行番号ずれ防止）
    for (let i = users.length - 1; i >= 0; i--) {
      if (String(users[i][COL_USER_ROLE - 1]) === '児童') {
        sheet.deleteRow(i + 2);
      }
    }
    return { success: true, message: '全児童を削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// =============================================
// パスワード管理機能 (教師用)
// =============================================

/**
 * 全児童のパスワード一覧を取得する
 * @returns {Object} - {success, passwords: [{seatNo, name, loginId, password}]}
 */
function getAllPasswords() {
  try {
    const users = getUserData();
    const passwords = users
      .filter(row => String(row[COL_USER_ROLE - 1]) === '児童')
      .map(row => ({
        seatNo: row[COL_USER_SEAT - 1],
        name: String(row[COL_USER_NAME - 1]),
        loginId: String(row[COL_USER_ID - 1]),
        password: String(row[COL_USER_PASS - 1])
      }))
      .sort((a, b) => Number(a.seatNo) - Number(b.seatNo));
    
    return { success: true, passwords };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message, passwords: [] };
  }
}

/**
 * パスワードを個別変更する
 * @param {string|number} seatNo - 出席番号
 * @param {string} newPassword - 新しいパスワード
 * @returns {Object} - {success, message}
 */
function updatePassword(seatNo, newPassword) {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    const targetIndex = users.findIndex(row => String(row[COL_USER_SEAT - 1]) === String(seatNo));
    
    if (targetIndex < 0) {
      return { success: false, message: '対象の児童が見つかりません。' };
    }
    
    const rowNum = targetIndex + 2;
    sheet.getRange(rowNum, COL_USER_PASS).setValue(String(newPassword).trim());
    return { success: true, message: 'パスワードを変更しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 個別パスワードをリセット（空にする）
 * @param {string|number} seatNo - 出席番号
 * @returns {Object} - {success, message}
 */
function resetPassword(seatNo) {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    const targetIndex = users.findIndex(row => String(row[COL_USER_SEAT - 1]) === String(seatNo));
    
    if (targetIndex < 0) {
      return { success: false, message: '対象の児童が見つかりません。' };
    }
    
    const rowNum = targetIndex + 2;
    sheet.getRange(rowNum, COL_USER_PASS).setValue('');
    return { success: true, message: 'パスワードをリセットしました。次回ログイン時に再設定が必要です。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 全児童のパスワードを一括リセット
 * @returns {Object} - {success, message}
 */
function resetAllPasswords() {
  try {
    const sheet = getUserSheet();
    const users = getUserData();
    users.forEach((row, i) => {
      if (String(row[COL_USER_ROLE - 1]) === '児童') {
        sheet.getRange(i + 2, COL_USER_PASS).setValue('');
      }
    });
    return { success: true, message: '全児童のパスワードをリセットしました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// =============================================
// 教師用: 児童ページ閲覧
// =============================================

/**
 * 特定児童の記録を取得する（教師閲覧用）
 * @param {string|number} seatNo - 出席番号
 * @returns {Object} - {success, records, studentName}
 */
function getStudentRecords(seatNo) {
  try {
    const users = getUserData();
    const studentRow = users.find(row => String(row[COL_USER_SEAT - 1]) === String(seatNo));
    const studentName = studentRow ? String(studentRow[COL_USER_NAME - 1]) : '不明';
    
    const result = getMyRecords(seatNo);
    return {
      success: result.success,
      records: result.records,
      studentName: studentName,
      message: result.message
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message, records: [] };
  }
}

// =============================================
// ユーティリティ
// =============================================

/**
 * 日付を yyyy/MM/dd 形式にフォーマット
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  try {
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy/MM/dd');
    }
    return String(dateValue).substring(0, 10);
  } catch (e) {
    return String(dateValue);
  }
}
