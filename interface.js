function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('🛠️ 檢查工具'); 
    menu.addItem('check TPR', 'compareandcheck');
    menu.addSeparator(); 
    menu.addItem('check CB', 'checkBilling');
    menu.addToUi(); 
}