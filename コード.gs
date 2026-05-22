/**
 * ウェブアプリにアクセスした際に実行される関数
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('林建DX：データ分析・インポート')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 分析用データを取得する
 */
function getAnalysisData(filters) {
  const ss = SpreadsheetApp.openById("1-HV-cb7tPiOvTD4nzKnlhfmPnfr59Kq2qxNNsSlblWU");
  
  let totalLabor = 0;
  let uniqueDays = new Set();
  let memberWorkData = {}; 
  let monthlyLabor = {}; 
  let categoryData = {}; 
  
  const importSheet = ss.getSheetByName("日報データ_取込");
  if (!importSheet) return null;
  
  const allData = importSheet.getDataRange().getValues();
  const reportRows = allData.slice(1);

  const today = new Date();
  const currentYm = `${today.getFullYear()}-${('0' + (today.getMonth() + 1)).slice(-2)}`;

  const trendMonths = [];
  if (filters.periodType === 'yearly') {
    const targetYear = filters.year || today.getFullYear();
    for (let m = 1; m <= 12; m++) {
      trendMonths.push(`${targetYear}-${('0' + m).slice(-2)}`);
    }
  } else {
    let baseYearMonth = filters.month || currentYm;
    const [bYear, bMonth] = baseYearMonth.split('-').map(Number);
    for (let i = 5; i >= 0; i--) {
      const d = new Date(bYear, bMonth - 1 - i, 1);
      trendMonths.push(`${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`);
    }
  }

  function isDateInRange(dateStr) {
    if (filters.periodType === 'all') return true;
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return false;
    const year = d.getFullYear();
    const month = ('0' + (d.getMonth() + 1)).slice(-2);
    const yyyymm = `${year}-${month}`;
    if (filters.periodType === 'monthly') return yyyymm === filters.month;
    if (filters.periodType === 'yearly') return year.toString() === filters.year;
    return true;
  }

  const allWorkTypes = new Set();

  reportRows.forEach(row => {
    const projNameInRow = row[31]; 
    const dateStr = row[1];        
    const labor = parseFloat(row[4]) || 0; 
    const member = row[6] || '不明';       
    const businessType = row[17] || '未分類事業'; 
    const workType = row[19] || '未分類作業';     

    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return;
    const yyyymm = `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`;

    // グラフ推移用の月別人役集計
    if (filters.project === 'all' || projNameInRow === filters.project) {
      monthlyLabor[yyyymm] = (monthlyLabor[yyyymm] || 0) + labor;
    }

    // フィルタリング
    if (filters.project !== 'all' && projNameInRow !== filters.project) return;
    if (!isDateInRange(dateStr)) return;
    
    totalLabor += labor;
    uniqueDays.add(dateStr.toString());

    if (!memberWorkData[member]) memberWorkData[member] = {};
    memberWorkData[member][workType] = (memberWorkData[member][workType] || 0) + labor;
    allWorkTypes.add(workType);

    if (!categoryData[businessType]) categoryData[businessType] = {};
    categoryData[businessType][workType] = (categoryData[businessType][workType] || 0) + labor;
  });

  // ========== 修正ポイント: 進捗管理シートを新しいファイルから取得する ==========
  const progressSs = SpreadsheetApp.openById("1hlzou-dX-hBmUwryu1YJgme4_IOiP4V8CswMQc3YXac");
  const progressSheet = progressSs.getSheetByName('施業管理アプリ - 進捗管理');
  
  let totalVolumeLog = 0;
  let totalVolumeChip = 0;
  let monthlyVolumeLog = {}; 
  let monthlyVolumeChip = {}; 
  
  let projectTableData = {}; // 表印刷用データ
  
  if (progressSheet) {
    const pData = progressSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      const row = pData[i];
      const projName = row[2];  // C列: 現場名
      if (!projName) continue;
      const targetMonth = row[3]; // D列: 対象月
      
      const volKanbatsu = parseFloat(row[12]) || 0; 
      const volKaibatsu = parseFloat(row[13]) || 0; 
      const volumeLog = volKanbatsu + volKaibatsu; 
      const volumeChip = parseFloat(row[14]) || 0; 

      let yyyymm = "";
      if (targetMonth instanceof Date) {
         yyyymm = `${targetMonth.getFullYear()}-${('0' + (targetMonth.getMonth() + 1)).slice(-2)}`;
      } else if (typeof targetMonth === 'string') {
         yyyymm = targetMonth.replace(/\//g, '-'); 
      }
      
      if (filters.project === 'all' || projName === filters.project) {
        monthlyVolumeLog[yyyymm] = (monthlyVolumeLog[yyyymm] || 0) + volumeLog;
        monthlyVolumeChip[yyyymm] = (monthlyVolumeChip[yyyymm] || 0) + volumeChip;
      }

      let inRange = false;
      let isPastOrCurrent = false;

      if (filters.periodType === 'all') {
         inRange = true;
         isPastOrCurrent = true;
      } else if (filters.periodType === 'monthly') {
         inRange = (yyyymm === filters.month);
         isPastOrCurrent = (yyyymm <= filters.month);
      } else if (filters.periodType === 'yearly') {
         inRange = yyyymm.startsWith(filters.year);
         isPastOrCurrent = (yyyymm <= filters.year + "-12");
      }

      // 表印刷用のデータ集約（フィルタ外の現場も累積計算のため保持、後で絞り込む）
      if (!projectTableData[projName]) {
        projectTableData[projName] = {
           name: projName, businessType: "", period: "", plannedArea: 0,
           doneAreaCumulative: 0, currentArea: 0, taxFree: "",
           volKanbatsu: 0, volKaibatsu: 0, volChip: 0, revenue: 0, hasDataInRange: false
        };
      }
      
      // 最新の固定情報を上書き（空白でなければ）
      if (row[5]) projectTableData[projName].period = row[5];
      if (row[8]) projectTableData[projName].plannedArea = parseFloat(row[8]);
      if (row[11]) projectTableData[projName].taxFree = row[11];

      // 累積面積（当月以前）
      if (isPastOrCurrent) {
        projectTableData[projName].doneAreaCumulative += (parseFloat(row[9]) || 0);
      }

      // 統計カード用（フィルタに合致する現場のみ）
      if (filters.project !== 'all' && projName !== filters.project) continue;
      
      if (inRange) {
        totalVolumeLog += volumeLog;
        totalVolumeChip += volumeChip;
        
        projectTableData[projName].hasDataInRange = true;
        projectTableData[projName].currentArea += (parseFloat(row[9]) || 0);
        projectTableData[projName].volKanbatsu += volKanbatsu;
        projectTableData[projName].volKaibatsu += volKaibatsu;
        projectTableData[projName].volChip += volumeChip;
        projectTableData[projName].revenue += ((parseFloat(row[19]) || 0) + (parseFloat(row[20]) || 0) + (parseFloat(row[21]) || 0));
      }
    }
  }

  // 日報から現場ごとの主力事業種を推定
  let projBizCount = {};
  reportRows.forEach(row => {
     const projName = row[31];
     const dateStr = row[1];
     if (filters.project !== 'all' && projName !== filters.project) return;
     if (!isDateInRange(dateStr)) return;
     const biz = row[17] || "";
     const labor = parseFloat(row[4]) || 0;
     if(biz) {
         if(!projBizCount[projName]) projBizCount[projName] = {};
         projBizCount[projName][biz] = (projBizCount[projName][biz] || 0) + labor;
     }
  });

  const tablePrintList = [];
  Object.keys(projectTableData).forEach(proj => {
     // 期間内にデータがある、または日報データが存在する現場のみ出力
     if (projectTableData[proj].hasDataInRange || projBizCount[proj]) {
         if(projBizCount[proj]) {
             // 人役が一番多い事業種を特定
             projectTableData[proj].businessType = Object.keys(projBizCount[proj]).reduce((a, b) => projBizCount[proj][a] > projBizCount[proj][b] ? a : b, "");
         }
         tablePrintList.push(projectTableData[proj]);
     }
  });

  const totalVolume = totalVolumeLog + totalVolumeChip;
  const averageProductivity = totalLabor > 0 ? (totalVolume / totalLabor).toFixed(1) : 0.0;

  // 2軸グラフ用データの生成
  const trendLabels = [];
  const trendDataProd = [];
  const trendDataVolLog = [];
  const trendDataVolChip = [];
  
  trendMonths.forEach(m => {
    trendLabels.push(m);
    const mLabor = monthlyLabor[m] || 0;
    const mVolLog = monthlyVolumeLog[m] || 0;
    const mVolChip = monthlyVolumeChip[m] || 0;
    const mVolTotal = mVolLog + mVolChip;
    
    trendDataVolLog.push(parseFloat(mVolLog.toFixed(1)));
    trendDataVolChip.push(parseFloat(mVolChip.toFixed(1)));
    
    if (m > currentYm) {
      trendDataProd.push(null);
    } else {
      const mProd = mLabor > 0 ? parseFloat((mVolTotal / mLabor).toFixed(1)) : 0;
      trendDataProd.push(mProd);
    }
  });

  const tableData = Object.keys(categoryData).map(b => {
    let bTotal = 0;
    let details = [];
    Object.entries(categoryData[b]).forEach(([w, val]) => {
      bTotal += val;
      details.push({ work: w, labor: parseFloat(val.toFixed(1)) });
    });
    details.sort((a, b) => b.labor - a.labor);
    return { business: b, total: parseFloat(bTotal.toFixed(1)), details: details };
  }).sort((a, b) => b.total - a.total);

  const sortedBusinessLabels = tableData.map(item => item.business);
  const workTypesSetForBiz = new Set();
  tableData.forEach(b => b.details.forEach(w => workTypesSetForBiz.add(w.work)));
  const workTypeArrayForBiz = Array.from(workTypesSetForBiz);

  const businessDatasets = workTypeArrayForBiz.map(wType => {
    const data = sortedBusinessLabels.map(bType => categoryData[bType][wType] || 0);
    return { label: wType, data: data };
  });

  const memberLabels = Object.keys(memberWorkData).sort();
  const memberDatasets = Array.from(allWorkTypes).map(wType => {
    return {
      label: wType,
      data: memberLabels.map(m => memberWorkData[m][wType] || 0)
    };
  });

  return {
    statLabor: totalLabor.toFixed(1),
    statDays: uniqueDays.size,
    statProductivity: averageProductivity,
    businessCategory: { labels: sortedBusinessLabels, datasets: businessDatasets, table: tableData },
    memberAnalysis: { labels: memberLabels, datasets: memberDatasets },
    productivityTrend: { labels: trendLabels, prodData: trendDataProd, volLogData: trendDataVolLog, volChipData: trendDataVolChip, periodType: filters.periodType },
    tablePrintData: tablePrintList
  };
}

/**
 * 現場一覧を取得する
 */
function getProjectList() {
  const ss = SpreadsheetApp.openById("1-HV-cb7tPiOvTD4nzKnlhfmPnfr59Kq2qxNNsSlblWU");
  const importSheet = ss.getSheetByName("日報データ_取込");
  if (!importSheet) return [];
  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) return [];
  const data = importSheet.getRange(2, 32, lastRow - 1, 1).getValues();
  return Array.from(new Set(data.map(row => row[0]).filter(name => name && name.toString().trim() !== "")));
}

/**
 * 精密パース関数 (GAS用)
 */
function parseTSVForGAS(text) {
  const result = [];
  let row = [];
  let col = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];
    if (char === '"') {
      if (inQuotes && nextChar === '"') { col += '"'; i++; }
      else { inQuotes = !inQuotes; }
      continue;
    }
    if (char === '\t' && !inQuotes) { row.push(col.trim()); col = ""; continue; }
    if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') i++;
      row.push(col.trim());
      if (row.length > 1 || (row.length === 1 && row[0] !== "")) result.push(row);
      row = []; col = ""; continue;
    }
    col += char;
  }
  if (col !== "" || row.length > 0) { row.push(col.trim()); result.push(row); }
  return result;
}

/**
 * 貼り付けられたデータを保存
 */
function importDailyReports(pasteData, targetProject) {
  try {
    const ss = SpreadsheetApp.openById("1-HV-cb7tPiOvTD4nzKnlhfmPnfr59Kq2qxNNsSlblWU");
    const sheet = ss.getSheetByName("日報データ_取込");
    if (!sheet) throw new Error("「日報データ_取込」シートが見つかりません。");

    const existingData = sheet.getDataRange().getValues();
    const existingKeys = new Set();
    if (existingData.length > 1) {
      existingData.slice(1).forEach(row => {
        const id = row[0];
        const proj = row[31];
        if (id) existingKeys.add(id + "_" + proj);
      });
    }

    const parsedRows = parseTSVForGAS(pasteData);
    const dataToAppend = [];
    let skipCount = 0;

    parsedRows.forEach(cols => {
      const reportId = cols[0];
      const dateStr = cols[1];
      if (!reportId || !String(reportId).match(/^\d+$/)) return;
      if (!dateStr || !String(dateStr).match(/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/)) return;

      const key = reportId + "_" + targetProject;
      if (existingKeys.has(key)) { skipCount++; return; }

      const newRow = new Array(32).fill("");
      cols.forEach((val, i) => { if (i < 31) newRow[i] = val; });
      newRow[31] = targetProject;
      dataToAppend.push(newRow);
    });

    if (dataToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, 32).setValues(dataToAppend);
    }
    return { success: true, total: parsedRows.length, count: dataToAppend.length, skipped: skipCount };
  } catch (e) {
    throw new Error(e.toString());
  }
}