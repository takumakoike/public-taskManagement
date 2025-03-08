function onOpen() :void {
    const customMenu = SpreadsheetApp.getUi();
    customMenu.createMenu('カスタム')
        .addItem('シート更新', 'createClientSheet') 
        .addToUi();
}