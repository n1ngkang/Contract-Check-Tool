// ==================== 全域常數 ====================
const COLORS = {
  ERROR: "#fff2cc",
  NOT_FOUND: "#ff9999",
  WARNING: "#c9daf8"
};

// 欄位定義：A=0, B=1, C=2...
const COL = {
  renewalMonth: 0, 
  link: 2,           // C 欄：合約連結
  id: 3,             // D 欄：合約 ID
  restaurantName: 4, // E 欄：店名
  type: 11,          // L 欄：Contract Type
  billingPeriod: 12, 
  sets: 13, 
  bonusSets: 14,     // O 欄：對應抽成 (Commission/ATM%)
  price: 15,         // P 欄：對應月費
  pricingType: 17,   // R 欄：B2B, B2C, Credit Card...
  overagePrice: 18   // S 欄：對應單價
};

const CSM = {
  id: 3, 
  billingPeriod: 38, 
  sets: 39, 
  bonusSets: 40, 
  price: 41,
  overagePrice: 44, 
  depositType: 46,       // AU 欄
  depositMonthlyFee: 47, // AV 欄
  commission: 48,        // AW 欄
  atmPercent: 49,        // AX 欄 (你的邏輯對應 AY)
  depositUnitPrice: 50,  // AY 欄 (你的邏輯對應 AW)
  voicePrice: 51, 
  surveyPrice: 52, 
  eCouponPrice: 53
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
    
    ERRORS = new Map();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, lastCol).setBackground(null);
    }
    
    const csmIndex = buildCSMIndexOptimized(data, CSM_SS);
    const contractMap = buildContractMap(data);
    const colorUpdates = new Map();
    const checked = new Set();
    
    let errorCount = 0;
    
    for (let r = 2; r < data.length; r++) {
      const row = data[r];
      const id = String(row[COL.id]).trim();
      const name = String(row[COL.restaurantName]).trim();
      const link = String(row[COL.link]).trim();
      const type = String(row[COL.type]).trim();
      const rowIndex = r + 1;
      
      if (!id) continue;
      
      const csmData = csmIndex[id];
      if (!csmData) {
        addError(r, name, rowIndex, id, "NOT_FOUND", "在 CSM 找不到對應資料", link, []);
        colorUpdates.set(`${r},${COL.id}`, COLORS.NOT_FOUND);
        errorCount++;
        continue;
      }
      
      const { csmRow } = csmData;
      
      if (!type) {
        addError(r, name, rowIndex, id, "WARNING", "Contract Type 為空", link, []);
        colorUpdates.set(`${r},${COL.type}`, COLORS.WARNING);
        continue;
      }
      
      let result;
      if (type === "Deposit" || type === "Deposit-Commission") {
        if (checked.has(id)) continue;
        checked.add(id);
        result = checkDeposit(csmRow, contractMap.get(id) || [], r, name, rowIndex, id, link, data);
      } else {
        result = checkContract(type, row, csmRow, r, name, rowIndex, id, link, data);
      }
      
      if (result) {
        applyResultsToMap(result, colorUpdates);
        if (result.hasError) errorCount++;
      }
    }
    
    batchUpdateColors(sheet, colorUpdates);
    showReport(TARGET_SS, "check TPR", data.length - 2, errorCount);
    Logger.log(`=== 總耗時: ${((new Date() - totalStart) / 1000).toFixed(1)}秒 ===`);
    
  } catch (e) {
    Logger.log('執行錯誤: ' + e.toString());
    SpreadsheetApp.getUi().alert('執行錯誤: ' + e.toString());
  }
}

// ==================== 邏輯函式 ====================

function checkContract(type, targetRow, csmRow, r, name, rowIndex, id, link) {
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
  
  if (result.errors.length > 0) addError(r, name, rowIndex, id, type, "數值不符", link, result.errors);
  return result;
}

function checkDeposit(csmRow, targetRows, r, name, rowIndex, id, link, data) {
  const result = { hasError: false, updates: [], errors: [] };
  const depositType = String(csmRow[CSM.depositType]).trim(); 
  
  const depositRow = targetRows.find(row => String(row[COL.type]).trim() === "Deposit");
  const commRows = targetRows.filter(row => String(row[COL.type]).trim() === "Deposit-Commission");

  // 1. B2B預付 / 綁卡
  if (depositType === "B2B 預付" || depositType === "B2B 綁卡") {
    if (!depositRow) {
      result.errors.push({ field: "Deposit", source: "缺少", target: "B2B行", type: "WARNING" });
    } else {
      const rowIdx = data.indexOf(depositRow);
      checkPricingType(depositRow, "B2B", result, rowIdx);
      checkBaseValues(depositRow, csmRow, result, true, rowIdx); 
    }
  }

  // 2. B2C預付 / 綁卡 / ATM
  else if (depositType === "B2C 預付" || depositType === "B2C 綁卡" || depositType === "B2C ATM") {
    if (!depositRow) {
      result.errors.push({ field: "Deposit", source: "缺少", target: "B2C行", type: "WARNING" });
    } else {
      const rowIdx = data.indexOf(depositRow);
      checkPricingType(depositRow, "B2C", result, rowIdx);
      checkBaseValues(depositRow, csmRow, result, false, rowIdx);
    }
    // 檢查抽成列
    commRows.forEach(cRow => {
      const rowIdx = data.indexOf(cRow);
      const pType = String(cRow[COL.pricingType]).trim();
      if (pType === "Credit Card") {
        checkBonusSetMatch(cRow, csmRow[CSM.commission], "Commission", result, rowIdx);
      } else if (pType === "Virtual ATM") {
        checkBonusSetMatch(cRow, csmRow[CSM.atmPercent], "ATM%", result, rowIdx);
      }
    });
  }

  // 3. B2B 轉帳
  else if (depositType === "B2B 轉帳") {
    if (depositRow) {
      const rowIdx = data.indexOf(depositRow);
      checkPricingType(depositRow, "B2B", result, rowIdx);
      checkBaseValues(depositRow, csmRow, result, false, rowIdx);
    }
    const atmRow = commRows.find(row => String(row[COL.pricingType]).trim() === "Virtual ATM");
    if (atmRow) {
      const rowIdx = data.indexOf(atmRow);
      checkBonusSetMatch(atmRow, csmRow[CSM.atmPercent], "ATM%", result, rowIdx);
    }
  }

  if (result.errors.length > 0) addError(r, name, rowIndex, id, "Deposit", "內容不符", link, result.errors);
  return result;
}

// ==================== 子檢查工具 ====================

// 檢查月費與單價
function checkBaseValues(row, csmRow, result, checkCommAsBonus, currentR) {
  // 月費標色
  const monthlyFeeErr = checkMonthlyFee(row, csmRow);
  if (monthlyFeeErr) {
    result.hasError = true;
    result.updates.push({ r: currentR, col: COL.price, color: COLORS.ERROR }); 
    result.errors.push({ field: "Monthly Fee", source: monthlyFeeErr.source, target: monthlyFeeErr.target, type: "ERROR" });
  }

  // 單價標色
  const csmUnit = norm(csmRow[CSM.depositUnitPrice]);
  const colUnit = norm(row[COL.overagePrice]);
  if (csmUnit !== "" && colUnit !== csmUnit) {
    result.hasError = true;
    result.updates.push({ r: currentR, col: COL.overagePrice, color: COLORS.ERROR }); 
    result.errors.push({ field: "Unit Price", source: colUnit, target: csmUnit, type: "ERROR" });
  }

  // B2B 的抽成 (在 Deposit 行)
  if (checkCommAsBonus) {
    checkBonusSetMatch(row, csmRow[CSM.commission], "Commission", result, currentR);
  }
}

// 檢查 Pricing Type (R 欄)
function checkPricingType(row, expected, result, currentR) {
  if (String(row[COL.pricingType]).trim() !== expected) {
    result.hasError = true;
    result.updates.push({ r: currentR, col: COL.pricingType, color: COLORS.ERROR }); 
    result.errors.push({ field: "Pricing Type", source: row[COL.pricingType], target: expected, type: "ERROR" });
  }
}

// 檢查 Bonus Sets (O 欄)
function checkBonusSetMatch(row, csmValue, label, result, currentR) {
  const target = norm(csmValue);
  const actual = norm(row[COL.bonusSets]);
  if (target !== "" && actual !== target) {
    result.hasError = true;
    result.updates.push({ r: currentR, col: COL.bonusSets, color: COLORS.ERROR }); // 這裡 push 標色
    result.errors.push({ field: label, source: actual, target: target, type: "ERROR" });
  }
}

// ==================== 輔助工具 ====================

function addError(r, name, rowIndex, id, type, msg, link, errors) {
  ERRORS.set(r, { 
    restaurantName: name, 
    rowIndex: rowIndex, 
    contractId: id, 
    contractUrl: link, 
    type: type, 
    message: msg, 
    errors: errors 
  });
}

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
      const id = String(rowData[0]).trim(); 
      if (id) {
        const correctedRow = new Array(60).fill("");
        correctedRow[CSM.id] = id;
        for (let colIdx = 38; colIdx <= 53; colIdx++) {
          correctedRow[colIdx] = rowData[colIdx - 3];
        }
        index[id] = { csmRow: correctedRow, sheetName: month };
      }
    });
  });
  return index;
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