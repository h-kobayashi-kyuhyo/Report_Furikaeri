function formatDataForDataPortal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // =========================================================
  // 【前半】回答データの成形（「集計用」シートの作成・更新）
  // =========================================================
  const sourceSheetName = 'フォームの回答 2';
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheetName = '集計用';
  let targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!sourceSheet) return console.error(`「${sourceSheetName}」シートが見つかりません。`);
  if (!targetSheet) targetSheet = ss.insertSheet(targetSheetName);
  
  // ★変更点1：サマリーシートと列名を完全に一致させ、「年月」も追加！
  let outputData = [];
  outputData.push([
    'タイムスタンプ', '年月', '報告日', '事業所', '部署', 'ヒヤリハットの有無', '区分', '内容', '要因', '評価(6段階)'
  ]);
  
  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  
  const colIdx = {
    timestamp: headers.indexOf('タイムスタンプ'),
    date: headers.indexOf('報告日'),
    location: headers.indexOf('所属事業所を選択してください'),
    role: headers.indexOf('担当を選択してください'),
    j1: headers.indexOf('J-1 ヒヤリとしたこと、ミスやトラブルが（些細なことでも）あった'),
    j2: headers.indexOf('J-2 次のどの区分に該当しそうですか？'),
    j3: headers.indexOf('J-3 簡単に内容を教えて'),
    j4: headers.findIndex(h => h && h.includes('J-4')),
    j5: headers.indexOf('J-5 ６段階で評価すると'),
    g1: headers.indexOf('G-1 ヒヤリとしたこと、ミスやトラブルが・・（些細なことでも）あった'),
    g2: headers.indexOf('G-2 次のどの区分に該当しそうですか？'),
    g3: headers.indexOf('G-3 簡単に内容を教えて'),
    g4: headers.findIndex(h => h && h.includes('G-4')),
    g5: headers.indexOf('G-5 ６段階で評価しましょう')
  };
  
  // 日付を「YYYY/MM」に変換する関数
  function getYearMonth(d) {
    if (!d) return null;
    let dateObj = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dateObj.getTime())) {
      if (typeof d === 'string' && d.match(/^\d{4}[\/-]\d{1,2}/)) {
        let parts = d.replace(/-/g, '/').split('/');
        return `${parts[0]}/${('0' + parts[1]).slice(-2)}`;
      }
      return null;
    }
    return `${dateObj.getFullYear()}/${('0' + (dateObj.getMonth() + 1)).slice(-2)}`;
  }

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (!row[colIdx.timestamp]) continue;
    
    let timestamp = row[colIdx.timestamp];
    let date = row[colIdx.date];
    let location = row[colIdx.location];
    let role = row[colIdx.role];
    let ym = getYearMonth(date); // ★「年月」を取得
    
    let hiyari = "", category = "", content = "", reason = "", rating = "";
    if (role === '事務') {
      hiyari = row[colIdx.j1]; category = row[colIdx.j2]; content = row[colIdx.j3];
      reason = colIdx.j4 !== -1 ? row[colIdx.j4] : ""; rating = row[colIdx.j5];
    } else if (role === '現場') {
      hiyari = row[colIdx.g1]; category = row[colIdx.g2]; content = row[colIdx.g3];
      reason = colIdx.g4 !== -1 ? row[colIdx.g4] : ""; rating = row[colIdx.g5];
    }
    
    // ★新しい列順で追加
    outputData.push([timestamp, ym, date, location, role, hiyari, category, content, reason, rating]);
  }
  
  targetSheet.clear();
  targetSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
  if (outputData.length > 1) {
    targetSheet.getRange(2, 1, outputData.length - 1, 1).setNumberFormat('yyyy/MM/dd hh:mm:ss'); // タイムスタンプ
    targetSheet.getRange(2, 3, outputData.length - 1, 1).setNumberFormat('yyyy/MM/dd'); // 報告日
  }

  // =========================================================
  // 【後半】返答率サマリーシートの作成
  // =========================================================
  const empSheet = ss.getSheetByName('社員リスト');
  if (!empSheet) return console.error('社員リストが見つかりません');
  const empData = empSheet.getDataRange().getValues();
  let empCounts = {}; 
  
  let eLocIdx = empData[0].indexOf('所属');
  let eDepIdx = empData[0].indexOf('部署');
  for (let i = 1; i < empData.length; i++) {
    let loc = empData[i][eLocIdx]; let dep = empData[i][eDepIdx];
    if (loc && dep) empCounts[`${loc}_${dep}`] = (empCounts[`${loc}_${dep}`] || 0) + 1;
  }

  // ★新しい配列の列番号に合わせてカウント
  let responseCounts = {};
  for (let i = 1; i < outputData.length; i++) {
    let row = outputData[i];
    let ym = row[1];  // 年月
    let loc = row[3]; // 事業所
    let dep = row[4]; // 部署
    
    if (ym && loc && dep) {
      let key = `${ym}_${loc}_${dep}`;
      responseCounts[key] = (responseCounts[key] || 0) + 1;
    }
  }

  const daysSheet = ss.getSheetByName('出勤日数');
  if (!daysSheet) return console.error('出勤日数シートが見つかりません');
  const daysData = daysSheet.getDataRange().getValues();
  
  let summaryData = [];
  summaryData.push(['年月', '事業所', '部署', '基準出勤日数', '社員数', '分母(日数×人数)', '分子(回答数)', '返答率']);

  let dDateIdx = daysData[0].indexOf('日付');
  let dLocIdx = daysData[0].indexOf('所属');
  let dDepIdx = daysData[0].indexOf('部署');
  let dDaysIdx = daysData[0].indexOf('日数');

  for (let i = 1; i < daysData.length; i++) {
    let dateVal = daysData[i][dDateIdx];
    let loc = daysData[i][dLocIdx];
    let dep = daysData[i][dDepIdx];
    let days = daysData[i][dDaysIdx];
    
    let ym = getYearMonth(dateVal);
    if (!ym || !loc || !dep) continue;

    let eCount = empCounts[`${loc}_${dep}`] || 0;
    let rCount = responseCounts[`${ym}_${loc}_${dep}`] || 0;
    let denominator = days * eCount;
    let rate = denominator > 0 ? (rCount / denominator) : 0;

    summaryData.push([ym, loc, dep, days, eCount, denominator, rCount, rate]);
  }

  const summarySheetName = '返答率サマリー';
  let summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) summarySheet = ss.insertSheet(summarySheetName);
  
  summarySheet.clear();
  summarySheet.getRange(1, 1, summaryData.length, summaryData[0].length).setValues(summaryData);
  if (summaryData.length > 1) summarySheet.getRange(2, 8, summaryData.length - 1, 1).setNumberFormat('0.0%');
}