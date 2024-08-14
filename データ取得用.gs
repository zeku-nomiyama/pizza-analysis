// データはスプシの各マスタから取得してくるようにする
// 出社日被り分析用メソッド
function get_count_for_co_working() {
  let sheet = get_spreadsheet_data('ピザマスタ');
  findCoWorkingDays(sheet);
}





// チーム被り分析用メソッド





// スプシデータ取得用メソッド
function get_spreadsheet_data(sheet_name) {
  const sheetId = '1-9OBj1J2KZ4HBY6LAfFvOntFzkvRIDLjOEtEz5QB9kU'
  let ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(sheet_name);
  return sheet;
}


function findCoWorkingDays(sheet) {
  // スプレッドシートとシートを取得
  const range = sheet.getRange('A2:B'); // 出社日がA列、社員名がB列にあると仮定
  const values = range.getValues();

  // 出社日ごとに社員名をグループ化
  let coWorkingMap = {};
  values.forEach(row => {
    const date = row[0];
    const employee = row[1];
    if (!coWorkingMap[date]) {
      coWorkingMap[date] = [];
    }
    coWorkingMap[date].push(employee);
  });

  // 社員ごとの出社日を比較して、同じ日に出社した回数をカウント
  let coWorkingCount = {};
  for (let date in coWorkingMap) {
    const employees = coWorkingMap[date];
    for (let i = 0; i < employees.length; i++) {
      for (let j = i + 1; j < employees.length; j++) {
        const pair = [employees[i], employees[j]].sort().join(' & ');
        if (!coWorkingCount[pair]) {
          coWorkingCount[pair] = 0;
        }
        coWorkingCount[pair]++;
      }
    }
  }

  // 結果を新しいシートに出力
  const resultSheet = get_spreadsheet_data("出社日被り分析用");
  resultSheet.clear(); // 既存のデータをクリア
  resultSheet.appendRow(['社員ペア', '同日出社回数']);
  for (let pair in coWorkingCount) {
    resultSheet.appendRow([pair, coWorkingCount[pair]]);
  }
}



