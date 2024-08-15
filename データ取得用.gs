// 社員ペアが同じ日に出社した回数をスプシに記録するメソッド
// 1度実行すると、既存のデータを削除して新たにデータ生成する
// すべて終わるまで中断しないこと
function recordEmployeePairAttendanceCount() {
  let sheet = getSpreadsheetData('ピザマスタ');
  findCoWorkingDays(sheet);
}

// 社員ペアがレクで同じチームになった回数をスプシに記録するメソッド
// 1度実行すると、既存のデータを削除して新たにデータ生成する
// すべて終わるまで中断しないこと
function recordEmployeePairTeamCount() {
  let sheet = getSpreadsheetData('ピザマスタ');
  countTeamOccurrences(sheet);
}



// 基本触らない
// スプシデータ取得用メソッド
function getSpreadsheetData(sheet_name) {
  const sheetId = '1-9OBj1J2KZ4HBY6LAfFvOntFzkvRIDLjOEtEz5QB9kU'
  let ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(sheet_name);
  return sheet;
}

// 出社日被り分析用メソッド
function findCoWorkingDays(pizzaSheet) {
  // 全社員の２人ペアの組み合わせを取得し、初期値を0とする
  const employeeSheet = getSpreadsheetData('社員マスタ');
  const employeeNameArray = employeeSheet.getRange('B2:B58').getValues().flat();
  let coWorkingCount = {};

  for (let i = 0; i < employeeNameArray.length; i++) {
    for (let j = i + 1; j < employeeNameArray.length; j++) {
      const pair = [employeeNameArray[i], employeeNameArray[j]].sort().join(' & ');
      coWorkingCount[pair] = 0;
    }
  }

  // スプレッドシートとシートを取得
  const range = pizzaSheet.getRange('A2:B'); // 出社日がA列、社員名がB列
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

  // 全社員のペアと出社日が同日のペアが一致する場合は、coWorkingCount[pair]を加算する。
  for (let date in coWorkingMap) {
    const employees = coWorkingMap[date];
    for (let i = 0; i < employees.length; i++) {
      for (let j = i + 1; j < employees.length; j++) {
        const pair = [employees[i], employees[j]].sort().join(' & ');
        if (coWorkingCount.hasOwnProperty(pair)) {
          coWorkingCount[pair]++;
        }
      }
    }
  }
  
  // 結果を新しいシートに出力
  const resultSheet = getSpreadsheetData("出社日被り分析用");
  resultSheet.clear(); // 既存のデータをクリア
  resultSheet.appendRow(['社員ペア', '同日出社回数']);
  for (let pair in coWorkingCount) {
    resultSheet.appendRow([pair, coWorkingCount[pair]]);
  }
}

// チーム被り分析用メソッド
function countTeamOccurrences(sheet) {
  // 全社員の２人ペアの組み合わせを取得し、初期値を0とする
  const employeeSheet = getSpreadsheetData('社員マスタ');
  const employeeNameArray = employeeSheet.getRange('B2:B58').getValues().flat();
  let pairCount = {};

  for (let i = 0; i < employeeNameArray.length; i++) {
    for (let j = i + 1; j < employeeNameArray.length; j++) {
      const pair = [employeeNameArray[i], employeeNameArray[j]].sort().join(' & ');
      pairCount[pair] = 0;
    }
  }

  // スプレッドシートとシートを取得
  const range = sheet.getRange('A2:E');
  const values = range.getValues();

  // 出社日ごとに社員名をグループ化
  let coWorkingMap = {};
  values.forEach(row => {
    const date = row[0];
    const employee = row[1];
    const team = row[4];
    if (!coWorkingMap[date]) {
      coWorkingMap[date] = [];
    }
    coWorkingMap[date].push([employee, team]);
  });

  // さらに同じチームごとに社員名をグループ化
  let teamMap = {};
  for (let date in coWorkingMap) {
    teamMap[date] = {};
    coWorkingMap[date].forEach(([employee, team]) => {
      if (!teamMap[date][team]) {
        teamMap[date][team] = [];
      }
      teamMap[date][team].push(employee);
    });
  }

  // 同じチーム内のすべての組み合わせを数える
  for (let date in teamMap) {
    for (let team in teamMap[date]) {
      const employees = teamMap[date][team];
      for (let i = 0; i < employees.length; i++) {
        for (let j = i + 1; j < employees.length; j++) {
          const pairKey = [employees[i], employees[j]].sort().join(' & ');
          if (pairCount.hasOwnProperty(pairKey)) {
            pairCount[pairKey]++;
          }
        }
      }
    }
  }

  // 結果をスプレッドシートに記録
  const resultSheet = getSpreadsheetData('チーム被り分析用'); 
  resultSheet.clear(); // 既存のデータをクリア
  resultSheet.appendRow(['社員ペア', 'チーム被り回数']);

  for (let pairKey in pairCount) {
    resultSheet.appendRow([pairKey, pairCount[pairKey]]);
  }
}
