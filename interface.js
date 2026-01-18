function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('🛠️ 檢查工具'); 
    menu.addItem('check TPR', 'compareandcheck'); 
    menu.addItem('check CB', 'checkBilling');
    menu.addItem('clear CB', 'clearCBSheet');
    menu.addToUi(); 
}