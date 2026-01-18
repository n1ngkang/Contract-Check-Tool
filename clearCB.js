/**
 * 清除 tool 工作表的指定輸入區域 (A3:D 以及 H3:K)
 * 改良版：精確鎖定 A 欄與 H 欄的最後一行
 */
function clearCBSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toolSheet = ss.getSheetByName("check CB");
  
  if (!toolSheet) {
    SpreadsheetApp.getUi().alert("找不到 'check CB' 工作表。");
    return;
  }

  const maxRows = toolSheet.getMaxRows();

  // 1. 分別找出 A 欄與 H 欄真正的最後一行 (往上跳)
  const lastRowA = toolSheet.getRange("A" + maxRows).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  // 檢查是否兩邊都沒資料
  if (lastRowA < 3) {
    SpreadsheetApp.getUi().alert("表單已經是空的了。");
    return;
  }

  // 二次確認
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認清除', '是否要清空工作表資料？', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    
    // 2. 清除 A 區 (如果 A3 有資料才清)
    if (lastRowA >= 3) {
      toolSheet.getRange("A3:DD" + lastRowA).clearContent();
    }
    
    ui.alert("✅ 資料已清除 ");
  }
}
