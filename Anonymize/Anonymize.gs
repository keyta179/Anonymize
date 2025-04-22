// 元データから個人情報を匿名化するプログラム
// メイン関数
function processStudentData() {
  const originalFile = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = originalFile.getSheets()[0]; // 先頭のシートを使用
  const data = inputSheet.getDataRange().getValues();
  const headers = data[0];

  const timestampIndex = headers.indexOf("タイムスタンプ");
  const gradeIndex = headers.indexOf("学年 / Grade");
  const studentIdIndex = headers.indexOf("学籍番号 / Student ID");
  const reasonIndex = headers.indexOf("来訪理由");

  const output = [["年", "月", "日", "時限", "学年", "学科", "クラス", "来訪理由"]];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const timestamp = row[timestampIndex];
    const grade = row[gradeIndex];
    const studentId = row[studentIdIndex];
    const reason = row[reasonIndex];

    const parsed = parseTimestampWithPeriod(timestamp); // timestamp は "2025/04/21 11:00:07(月2)" など
    const department = getDepartment(studentId);
    const group = getGroup(studentId);

    output.push([parsed.year, parsed.month, parsed.date, parsed.period, grade, department, group, reason]);
  }

  const newSpreadsheet = SpreadsheetApp.create("過去データ_" + new Date().toISOString().slice(0, 10));
  const newSheet = newSpreadsheet.getActiveSheet();
  newSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  const originalFileId = originalFile.getId();
  const newFileId = newSpreadsheet.getId();
  const originalFileParents = DriveApp.getFileById(originalFileId).getParents();
  while (originalFileParents.hasNext()) {
    const parent = originalFileParents.next();
    parent.addFile(DriveApp.getFileById(newFileId));
  }
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(newFileId));
}


  // 「年・月・日・曜日＋時限」をすべて返す関数
 function parseTimestampWithPeriod(timestampStr) {
  // "2025/04/21 11:00:07(月2)" の形式から年月日と曜日・時限を抽出
  const match = timestampStr.match(/^(\d{4})\/(\d{2})\/(\d{2}) [0-9:]+(?:\(([月火水木金土日])([2L345]?)\))?$/);

  if (!match) {
    console.warn("形式エラー: " + timestampStr);
    return {
      year: "", month: "", date: "", period: ""
    };
  }

  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const date = parseInt(match[3], 10);
  const day = match[4] || "";
  const period = match[5] || "";

  return {
    year: year,
    month: month,
    date: date,
    period: day + period // 例："月2"
  };
}


// 学科を返す関数
function getDepartment(studentId) {
  const upperId = studentId.toUpperCase();
  if (upperId.length > 3 && upperId[2] === "K") {
    if (upperId[3] === "0") return "CS";
    if (upperId[3] === "1") return "DM";
  }
  return "others";
}

// クラスを返す関数
function getGroup(studentId) {
  const upperId = studentId.toUpperCase();
  if (upperId.length > 4 && upperId[2] === "K") {
    const code = upperId.substring(3, 5);
    if (code === "00") return "A";
    if (code === "01") return "B";
    if (code === "10") return "C";
    if (code === "11") return "D";
  }
  return "others";
}