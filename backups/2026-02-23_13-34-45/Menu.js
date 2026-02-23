function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Documentos')
    .addItem('Generar Contratos', 'generarContratos')   
    .addItem('Archivar Contratos', 'organizarContratosFirmados')   
    
    .addToUi();
}