// ==================== 全域常數 ====================
const COLORS = {
  ERROR: "#fff2cc",
  NOT_FOUND: "#ff9999",
  WARNING: "#c9daf8"
};

const COL = {
  renewalMonth: 0, id: 3, restaurantName: 4, type: 11,
  billingPeriod: 12, sets: 13, bonusSets: 14, price: 15,
  pricingType: 17, overagePrice: 18
};

const CSM = {
  id: 3, billingPeriod: 38, sets: 39, bonusSets: 40, price: 41,
  overagePrice: 44, depositType: 46, depositMonthlyFee: 47,
  commission: 48, atmPercent: 49, depositUnitPrice: 50,
  voicePrice: 51, surveyPrice: 52, eCouponPrice: 53
};

const RULES = {
  RSV: [
    { csm: CSM.billingPeriod, col: COL.billingPeriod, name: "Billing Period" },
    { csm: CSM.sets, col: COL.sets, name: "Sets" },
    { csm: CSM.bonusSets, col: COL.bonusSets, name: "Bonus Sets" },
    { csm: CSM.price, col: COL.price, name: "Price" },
    { csm: CSM.overagePrice, col: COL.overagePrice, name: "Overage Price" }
  ],
  RwG: [
    { csm: CSM.billingPeriod, col: COL.billingPeriod, name: "Billing Period" },
    { csm: CSM.sets, col: COL.sets, name: "Sets" },
    { csm: CSM.bonusSets, col: COL.bonusSets, name: "Bonus Sets" },
    { csm: CSM.price, col: COL.price, name: "Price" },
    { csm: CSM.overagePrice, col: COL.overagePrice, name: "Overage Price" }
  ],
  Voice: [
    { csm: CSM.billingPeriod, col: COL.billingPeriod, name: "Billing Period" },
    { csm: CSM.voicePrice, col: COL.price, name: "Price" }
  ],
  Survey: [
    { csm: CSM.billingPeriod, col: COL.billingPeriod, name: "Billing Period" },
    { csm: CSM.surveyPrice, col: COL.price, name: "Price" }
  ],
  eCoupon: [
    { csm: CSM.billingPeriod, col: COL.billingPeriod, name: "Billing Period" },
    { csm: CSM.eCouponPrice, col: COL.price, name: "Price" }
  ]
};

let ERRORS = new Map();

// ==================== 主程式 ====================
function compareandcheck() {
  let TARGET_SS;
  try {
    const totalStart = new Date();
    Logger.log('=== 開始執行 ===');
    
    TARGET_SS = SpreadsheetApp.getActiveSpreadsheet();
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const csmId = scriptProperties.getProperty('CSM_FILE_ID');
    if (!csmId) {
      throw new Error("找不到 CSM_FILE_ID 屬性，請先至設定中新增。");
    }
    const CSM_SS = SpreadsheetApp.openById(csmId);
    
    const sheet = TARGET_SS.getSheetByName("check TPR");
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('錯誤', '找不到工作表 "check TPR"', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    Logger.log(`開啟檔案完成: ${((new Date() - totalStart) / 1000).toFixed(1)}秒`);
    
    ERRORS = new Map();
    
    // 讀取本地資料
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    // 清空底色
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, lastCol).setBackground(null);
    }
    
    // 建立 CSM 索引 (改用區段讀取，避開大量公式)
    const csmStart = new Date();
    const csmIndex = buildCSMIndexOptimized(data, CSM_SS);
    Logger.log(`建立 CSM 索引: ${((new Date() - csmStart) / 1000).toFixed(1)}秒`);
    
    const contractMap = buildContractMap(data);
    const colorUpdates = new Map();
    const checked = new Set();
    
    let errorCount = 0;
    const checkStart = new Date();
    
    for (let r = 2; r < data.length; r++) {
      const row = data[r];
      const id = String(row[COL.id]).trim();
      const name = String(row[COL.restaurantName]).trim();
      const type = String(row[COL.type]).trim();
      const rowIndex = r + 1;
      
      if (!id) continue;
      
      const csmData = csmIndex[id];
      if (!csmData) {
        addError(r, name, rowIndex, id, "NOT_FOUND", "在 CSM 找不到對應資料", []);
        colorUpdates.set(`${r},${COL.id}`, COLORS.NOT_FOUND);
        errorCount++;
        continue;
      }
      
      const { csmRow, sheetName } = csmData;
      
      if (!type) {
        addError(r, name, rowIndex, id, "WARNING", "Contract Type 為空", []);
        colorUpdates.set(`${r},${COL.type}`, COLORS.WARNING);
        continue;
      }
      
      // 檢查邏輯
      let result;
      if (type === "Deposit" || type === "Deposit-Commission") {
        if (checked.has(id)) continue;
        checked.add(id);
        result = checkDeposit(csmRow, contractMap.get(id) || [], r, name, rowIndex, id, sheetName);
      } else {
        result = checkContract(type, row, csmRow, r, name, rowIndex, id, sheetName);
      }
      
      if (result) {
        applyResultsToMap(result, colorUpdates);
        if (result.hasError) errorCount++;
      }
    }
    
    Logger.log(`檢查邏輯完成: ${((new Date() - checkStart) / 1000).toFixed(1)}秒`);
    
    batchUpdateColors(sheet, colorUpdates);
    
    // 顯示報告
    showReport(TARGET_SS, "check TPR", data.length - 2, errorCount);
    Logger.log(`=== 總耗時: ${((new Date() - totalStart) / 1000).toFixed(1)}秒 ===`);
    
  } catch (e) {
    Logger.log('執行錯誤: ' + e.toString());
    SpreadsheetApp.getUi().alert('執行錯誤: ' + e.toString());
  }
}

// ==================== 讀取函式 ====================

function buildCSMIndexOptimized(data, csmSS) {
  const index = {};
  const months = [...new Set(data.slice(2).map(r => parseMonth(r[COL.renewalMonth])).filter(Boolean))];
  
  months.forEach(month => {
    const sheet = csmSS.getSheetByName(month);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const fullRangeData = sheet.getRange(2, 4, lastRow - 1, 51).getValues(); 

    fullRangeData.forEach(rowData => {
      const id = String(rowData[0]).trim(); // rowData[0] 是 D 欄
      if (id) {
        const correctedRow = new Array(55).fill("");
        correctedRow[CSM.id] = id;
        
        // 將 rowData (從 index 0 開始) 填回原本的列位置 (例如 38 欄在 rowData 裡是 index 35)
        for (let colIdx = 38; colIdx <= 53; colIdx++) {
          correctedRow[colIdx] = rowData[colIdx - 3];
        }
        index[id] = { csmRow: correctedRow, sheetName: month };
      }
    });
  });
  return index;
}

// ==================== 顏色更新 ====================

function applyResultsToMap(result, colorUpdates) {
  result.updates.forEach(u => {
    const key = `${u.r},${u.col}`;
    const existing = colorUpdates.get(key);
    if (existing === COLORS.NOT_FOUND) return;
    if (existing === COLORS.ERROR && u.color === COLORS.WARNING) return;
    colorUpdates.set(key, u.color);
  });
}

function batchUpdateColors(sheet, colorUpdates) {
  if (colorUpdates.size === 0) return;
  
  const colorGroups = new Map();
  colorUpdates.forEach((color, key) => {
    if (!colorGroups.has(color)) colorGroups.set(color, []);
    const [r, c] = key.split(',').map(Number);
    colorGroups.get(color).push(sheet.getRange(r + 1, c + 1).getA1Notation());
  });
  
  colorGroups.forEach((a1Notations, color) => {
    const batchSize = 100;
    for (let i = 0; i < a1Notations.length; i += batchSize) {
      const batch = a1Notations.slice(i, i + batchSize);
      sheet.getRangeList(batch).setBackground(color);
    }
  });
}

// ==================== 輔助邏輯 ====================

function checkContract(type, targetRow, csmRow, r, name, rowIndex, id, sheetName) {
  const rules = RULES[type];
  if (!rules) return { hasError: false, updates: [], errors: [] };
  const result = { hasError: false, updates: [], errors: [] };
  
  rules.forEach(rule => {
    const csmVal = norm(csmRow[rule.csm]);
    const colVal = norm(targetRow[rule.col]);
    
    if ((type === "Voice" || type === "Survey" || type === "eCoupon") && rule.name === "Price") {
      const monthlyCheck = checkServiceMonthlyFee(targetRow, csmRow, rule.csm);
      if (monthlyCheck) {
        if (monthlyCheck.type === "ERROR") {
          result.hasError = true;
          result.updates.push({ r, col: rule.col, color: COLORS.ERROR });
          result.errors.push({ field: rule.name, source: monthlyCheck.source, target: monthlyCheck.target, type: "ERROR" });
        } else if (monthlyCheck.type === "WARNING") {
          result.updates.push({ r, col: rule.col, color: COLORS.WARNING });
          result.errors.push({ field: rule.name, source: monthlyCheck.source, target: "", type: "WARNING" });
        }
      }
      return;
    }
    
    if (csmVal !== "" && csmVal !== colVal) {
      result.hasError = true;
      result.updates.push({ r, col: rule.col, color: COLORS.ERROR });
      result.errors.push({ field: rule.name, source: colVal, target: csmVal, type: "ERROR" });
    } else if (csmVal === "" && colVal !== "") {
      result.updates.push({ r, col: rule.col, color: COLORS.WARNING });
      result.errors.push({ field: rule.name, source: colVal, target: "", type: "WARNING" });
    }
  });
  if (result.errors.length > 0) addError(r, name, rowIndex, id, type, "數值不符", result.errors);
  return result;
}

function checkDeposit(csmRow, targetRows, r, name, rowIndex, id, sheetName) {
  const result = { hasError: false, updates: [], errors: [] };
  const depositType = String(csmRow[CSM.depositType]).trim();
  const depositRow = targetRows.find(row => String(row[COL.type]).trim() === "Deposit");
  const commRows = targetRows.filter(row => String(row[COL.type]).trim() === "Deposit-Commission");
  
  if (!depositType || depositType === "N/A") {
    if (depositRow) {
      result.updates.push({ r, col: COL.type, color: COLORS.WARNING });
      result.errors.push({ field: "Deposit Type", source: "Deposit", target: "", type: "WARNING" });
    }
    if (result.errors.length > 0) addError(r, name, rowIndex, id, "Deposit", "CSM 無紀錄", result.errors);
    return result;
  }
  
  if (!depositRow) {
    result.errors.push({ field: "Deposit", source: "", target: depositType, type: "WARNING" });
    addError(r, name, rowIndex, id, "Deposit", "缺少 Deposit 合約", result.errors);
    return result;
  }
  
  const baseChecks = [
    { csm: CSM.billingPeriod, col: COL.billingPeriod, name: "Billing Period" },
    { csm: CSM.depositUnitPrice, col: COL.overagePrice, name: "Unit Price" }
  ];
  
  baseChecks.forEach(rule => {
    const csmVal = norm(csmRow[rule.csm]);
    const colVal = norm(depositRow[rule.col]);
    if (csmVal !== "" && csmVal !== colVal) {
      result.hasError = true;
      result.updates.push({ r, col: rule.col, color: COLORS.ERROR });
      result.errors.push({ field: rule.name, source: colVal, target: csmVal, type: "ERROR" });
    }
  });
  
  const monthlyFee = checkMonthlyFee(depositRow, csmRow);
  if (monthlyFee) {
    result.hasError = true;
    result.updates.push({ r, col: COL.price, color: COLORS.ERROR });
    result.errors.push({ field: "Monthly Fee", source: monthlyFee.source, target: monthlyFee.target, type: "ERROR" });
  }
  
  if (result.errors.length > 0) addError(r, name, rowIndex, id, "Deposit", "Deposit 錯誤", result.errors);
  return result;
}

function checkMonthlyFee(depositRow, csmRow) {
  const period = String(depositRow[COL.billingPeriod]).trim().toLowerCase();
  const source = Number(depositRow[COL.price]);
  const target = Number(csmRow[CSM.depositMonthlyFee]);
  if (!target) return null;
  const expected = period === "annual" ? Math.round(source / 12 * 100) / 100 : Math.round(source * 100) / 100;
  const actual = Math.round(target * 100) / 100;
  return expected !== actual ? { source: String(expected), target: String(actual) } : null;
}

function checkServiceMonthlyFee(targetRow, csmRow, csmPriceCol) {
  const period = String(targetRow[COL.billingPeriod]).trim().toLowerCase();
  const sourcePrice = Number(targetRow[COL.price]);
  const targetPrice = Number(csmRow[csmPriceCol]);
  if (String(csmRow[csmPriceCol]).trim() === "") return sourcePrice > 0 ? { source: String(sourcePrice), target: "", type: "WARNING" } : null;
  const expected = period === "annual" ? Math.round(sourcePrice / 12 * 100) / 100 : Math.round(sourcePrice * 100) / 100;
  const actual = Math.round(targetPrice * 100) / 100;
  return expected !== actual ? { source: String(expected), target: String(actual), type: "ERROR" } : null;
}

function norm(val) {
  if (val === null || val === undefined || val === "") return "";
  const str = String(val).trim();
  const cleaned = str.replace(/[\s,$¥£€％%\u00A0]/g, "").replace(/−/g, "-");
  const num = Number(cleaned);
  return !isNaN(num) ? String(Math.round(num * 1e6) / 1e6) : str.toLowerCase();
}

function parseMonth(input) {
  if (!input) return null;
  const str = String(input).trim();
  const m = str.match(/(\d{4})\/(\d{1,2})/);
  return m ? `${m[1]}${m[2].padStart(2, '0')}` : (/^\d{6}$/.test(str) ? str : null);
}

function addError(r, name, rowIndex, id, type, msg, errors) {
  ERRORS.set(r, { restaurantName: name, rowIndex, contractId: id, type, message: msg, errors });
}

function buildContractMap(data) {
  const map = new Map();
  for (let r = 2; r < data.length; r++) {
    const id = String(data[r][COL.id]).trim();
    if (id) {
      if (!map.has(id)) map.set(id, []);
      map.get(id).push(data[r]);
    }
  }
  return map;
}

function showReport(ss, sheetName, total, errors) {
  try {
    const template = HtmlService.createTemplateFromFile('ReportDialog');
    template.reportData = Array.from(ERRORS.values());
    template.totalChecked = total;
    template.errorCount = errors;
    template.fileId = ss.getId();
    template.sheetName = sheetName;
    const htmlOutput = template.evaluate().setTitle('檢查報告').setWidth(350);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (e) {
    Logger.log('顯示報告時發生錯誤: ' + e.toString());
  }
}