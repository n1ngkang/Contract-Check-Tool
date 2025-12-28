const CONFIG = {
  BILLING_SHEET_NAME: "check CB",
  PLAN_SHEET_NAME: "方案表",
  PLAN_RANGE_RSV: "A1:C28",    
  PLAN_RANGE_DEPOSIT: "E1:H27", 
  
  COL_START_ROW: 3,
  COL_BILLING_PERIOD: 33, // AG
  COL_RSV_SETS: 36,       // AJ
  COL_PRICE: 38,          // AL
  COL_SHARED_ID: 40,      // AN
  COL_OVERAGE: 44,        // AR
  
  COLOR_NORMAL: null,          
  COLOR_NEGOTIATED: "#e6f3ff", // 淡藍色 (議價合理)
  COLOR_ERROR: "#fff2cc",      // 黃色 (明顯錯誤)
};

const COL_DEPOSIT = {
  MONTHLY: 76, // BX
  ANNUAL: 77,  // BY
  UNIT_PRICE: 79, // CA
  B2B_COMM: 80,   // CB
  B2C_COMM: 81    // CC
};

function checkBilling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const billingSheet = ss.getSheetByName(CONFIG.BILLING_SHEET_NAME);
  const planSheet = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);

  if (!billingSheet || !planSheet) {
    SpreadsheetApp.getUi().alert("找不到工作表，請檢查名稱是否正確。");
    return;
  }

  // --- 1. 預處理索引 ---
  const planDataRsv = planSheet.getRange(CONFIG.PLAN_RANGE_RSV).getValues();
  const { planIndex, rangeIndex, overageIndex } = buildEnhancedPlanIndex(planDataRsv);
  
  const planDataDeposit = planSheet.getRange(CONFIG.PLAN_RANGE_DEPOSIT).getValues();
  const depositIndex = buildDepositPlanIndex(planDataDeposit);

  const billingData = billingSheet.getDataRange().getValues();
  const sharedIdMap = buildSharedIdMap(billingData);
  
  const fullRange = billingSheet.getDataRange();
  const colorMatrix = fullRange.getBackgrounds();

  let errorCount = 0;

  // --- 2. 檢查迴圈 ---
  for (let i = CONFIG.COL_START_ROW - 1; i < billingData.length; i++) {
    const row = billingData[i];
    const rowIndex = i;

    // A. 啟動即清空舊顏色 (只針對檢查範圍)
    const resetCols = [
      CONFIG.COL_PRICE, CONFIG.COL_OVERAGE, 
      COL_DEPOSIT.MONTHLY, COL_DEPOSIT.ANNUAL, COL_DEPOSIT.UNIT_PRICE
    ];
    resetCols.forEach(col => colorMatrix[rowIndex][col - 1] = CONFIG.COLOR_NORMAL);

    const hasRsv = row[CONFIG.COL_RSV_SETS - 1] !== "";
    const hasDeposit = (Number(row[COL_DEPOSIT.MONTHLY - 1]) || 0) > 0 || (Number(row[COL_DEPOSIT.ANNUAL - 1]) || 0) > 0;
    
    if (!hasRsv && !hasDeposit) continue;

    // B. RSV 精準檢查
    if (hasRsv) {
      const rsvStatus = getRsvPreciseStatus(row, planIndex, rangeIndex, overageIndex, sharedIdMap);
      colorMatrix[rowIndex][CONFIG.COL_PRICE - 1] = rsvStatus.priceColor;
      colorMatrix[rowIndex][CONFIG.COL_OVERAGE - 1] = rsvStatus.overageColor;
      if (rsvStatus.priceColor === CONFIG.COLOR_ERROR || rsvStatus.overageColor === CONFIG.COLOR_ERROR) errorCount++;
    }

    // C. Deposit 精準檢查 (包含 B2B 預設邏輯)
    if (hasDeposit) {
      const depStatus = getDepositPreciseStatus(row, depositIndex);
      colorMatrix[rowIndex][COL_DEPOSIT.MONTHLY - 1] = depStatus.feeColor;
      colorMatrix[rowIndex][COL_DEPOSIT.ANNUAL - 1] = depStatus.feeColor;
      colorMatrix[rowIndex][COL_DEPOSIT.UNIT_PRICE - 1] = depStatus.unitPriceColor;
      if (depStatus.feeColor === CONFIG.COLOR_ERROR || depStatus.unitPriceColor === CONFIG.COLOR_ERROR) errorCount++;
    }
  }

  // --- 3. 一次性寫入 ---
  fullRange.setBackgrounds(colorMatrix);
  SpreadsheetApp.flush();
  ss.toast(`檢查完成!共發現 ${errorCount} 行異常`, '✅ 完成', 3);
}

// ====================================================
// 📌 核心邏輯：RSV (Shared ID & 區間判定)
// ====================================================

function getRsvPreciseStatus(row, planIndex, rangeIndex, overageIndex, sharedIdMap) {
  const period = String(row[CONFIG.COL_BILLING_PERIOD - 1]).toLowerCase();
  let sets = String(row[CONFIG.COL_RSV_SETS - 1]).toLowerCase();
  let price = Number(row[CONFIG.COL_PRICE - 1]) || 0;
  const overage = Number(row[CONFIG.COL_OVERAGE - 1]) || 0;
  const sharedId = String(row[CONFIG.COL_SHARED_ID - 1]).trim();

  if (period === 'annual') {
    if (!isNaN(sets)) sets = String(Number(sets) / 12);
    price = price / 12;
  }

  let status = { priceColor: CONFIG.COLOR_NORMAL, overageColor: CONFIG.COLOR_NORMAL };

  // Overage 判定 (Shared ID 優先於方案表)
  const isSharedOverage = (sharedId !== "" && sharedIdMap[sharedId] === overage);
  const isStandardOverage = overageIndex.has(overage);

  if (!isSharedOverage && !isStandardOverage) {
    status.overageColor = CONFIG.COLOR_ERROR;
  }

  // Price 判定
  const roundedPrice = Math.round(price * 100) / 100;
  const lookupKey = `${sets}|${roundedPrice}|${overage}`;
  const range = rangeIndex[sets];

  if (planIndex.has(lookupKey)) {
    status.priceColor = CONFIG.COLOR_NORMAL;
  } else if (range && price >= range.min && price <= range.max) {
    status.priceColor = CONFIG.COLOR_NEGOTIATED;
  } else {
    status.priceColor = CONFIG.COLOR_ERROR;
  }

  return status;
}

// ====================================================
// 📌 核心邏輯：Deposit (無佣金預設 B2B)
// ====================================================

function getDepositPreciseStatus(row, depositIndex) {
  const annual = Number(row[COL_DEPOSIT.ANNUAL - 1]) || 0;
  const monthly = Number(row[COL_DEPOSIT.MONTHLY - 1]) || 0;
  const unitPrice = Number(row[COL_DEPOSIT.UNIT_PRICE - 1]) || 0;
  const b2b = Number(row[COL_DEPOSIT.B2B_COMM - 1]) || 0;
  const b2c = Number(row[COL_DEPOSIT.B2C_COMM - 1]) || 0;

  const currentMonthly = annual > 0 ? Math.round((annual / 12) * 100) / 100 : Math.round(monthly * 100) / 100;
  
  // 優化判斷：沒有 B2C 佣金就一律當作 B2B
  let currentType = "B2B";
  let currentComm = b2b;
  if (b2c > 0) {
    currentType = "B2C";
    currentComm = b2c;
  }

  let status = { feeColor: CONFIG.COLOR_ERROR, unitPriceColor: CONFIG.COLOR_ERROR };

  for (let [key, _] of depositIndex) {
    const [pType, pPrice, pUp, pComm] = key.split('|');
    
    const isTypeMatch = (pType === currentType);
    const isCommMatch = Math.abs(Number(pComm) - currentComm) < 0.0001;

    if (isTypeMatch && isCommMatch) {
      if (Math.abs(Number(pPrice) - currentMonthly) < 0.1) status.feeColor = CONFIG.COLOR_NORMAL;
      if (Math.abs(Number(pUp) - unitPrice) < 0.1) status.unitPriceColor = CONFIG.COLOR_NORMAL;
    }
  }

  return status;
}

// ====================================================
// 📌 輔助函式 (索引建立)
// ====================================================

function buildEnhancedPlanIndex(planData) {
  const planIndex = new Map();
  const rangeIndex = {};
  const overageIndex = new Set();

  planData.forEach(row => {
    const sets = String(row[0]).trim().toLowerCase();
    const price = Number(row[1]);
    const overage = Number(row[2]);
    if (!sets || sets === "sets") return;

    planIndex.set(`${sets}|${Math.round(price * 100) / 100}|${overage}`, true);
    overageIndex.add(overage);

    if (!isNaN(price) && price > 0) {
      if (!rangeIndex[sets]) {
        rangeIndex[sets] = { min: price, max: price };
      } else {
        rangeIndex[sets].min = Math.min(rangeIndex[sets].min, price);
        rangeIndex[sets].max = Math.max(rangeIndex[sets].max, price);
      }
    }
  });
  return { planIndex, rangeIndex, overageIndex };
}

function buildDepositPlanIndex(data) {
  const index = new Map();
  data.forEach(row => {
    const type = String(row[0]).trim().toUpperCase();
    const price = Number(row[1]) || 0;
    const up = Number(row[2]) || 0;
    const comm = Number(row[3]) || 0;
    if (type !== "") {
      index.set(`${type}|${Math.round(price * 100) / 100}|${up}|${comm}`, true);
    }
  });
  return index;
}

function buildSharedIdMap(data) {
  const map = {};
  for (let i = CONFIG.COL_START_ROW - 1; i < data.length; i++) {
    const id = String(data[i][CONFIG.COL_SHARED_ID - 1]).trim();
    const ov = Number(data[i][CONFIG.COL_OVERAGE - 1]);
    if (id && !isNaN(ov)) map[id] = Math.max(map[id] || 0, ov);
  }
  return map;
}