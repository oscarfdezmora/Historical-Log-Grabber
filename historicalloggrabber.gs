function lastBlankRow(link,tab){
// Dado un link y una pestaña, devuelve la primera línea en completamente en blanco de un documento;
  var values = SpreadsheetApp.openByUrl(link).getSheetByName(tab).getDataRange().getValues();
  var row = 0;
  for (var row=0; row<values.length; row++) {
    if (!values[row].join("")) break;
  }
    return row+1;
}

function copyHereWithDate(link,tab){
// Copiar a la hoja activa 
  var thisSheet = SpreadsheetApp.getActive();
  var tabName = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd") + " " + tab;
  var originSheet = SpreadsheetApp.openByUrl(link);
   
// Copyto si se utiliza con Sheet directamente se trae las formulas y datos fijado; no vale
// Establecidos los limites en a1:ez3000... hay que revisar si nos pasamos de esto periódicamente.
// Se pone límite en hoja porque al crearse, el valor de filas es 1k y casca
  var datos = originSheet.getSheetByName(tab).getRange("a1:ez3000").getValues();
  SpreadsheetApp.getActive().insertSheet(tabName).getRange("a1:ez3000").setValues(datos);

}

function importHeaders(link,origintab,destinytab){
// Trae las cabeceras actualizadas con la información
  var datos = SpreadsheetApp.openByUrl(link).getSheetByName(origintab).getRange("a:fz").getValues();
  SpreadsheetApp.getActive().getSheetByName(destinytab).getRange("a:fz").setValues(datos);
}

function importData(link,origintab,destinytab){
// Trae los datos actualizados de las variables y paises

// Fecha en la que se ejecuta
  var date = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  var destinyUrl = SpreadsheetApp.getActive().getUrl();
  var lastDataOrigin = lastBlankRow(link,origintab) - 1; // -1 para corrección del dato
  var firstRowDataDestiny = lastBlankRow(destinyUrl,destinytab);
  var lastRowDataDestiny = firstRowDataDestiny + lastDataOrigin - 5 - 1; // -5 por la cabecera del Origen, -1 de corrección
  
// Pone fecha en los datos en la primera columna
  SpreadsheetApp.getActive().getSheetByName(destinytab).getRange(firstRowDataDestiny,1,lastDataOrigin - 5).setValue(date);
  
  var originRange = "c6:fz" + lastDataOrigin; // Define rango de datos de collector
  var destinyRange = "b" + firstRowDataDestiny + ":fy" + lastRowDataDestiny; // Define rango de datos en destino
  
// Copia el dato de origentab y pega en destino
  var data = SpreadsheetApp.openByUrl(link).getSheetByName(origintab).getRange(originRange).getValues();
  SpreadsheetApp.getActive().getSheetByName(destinytab).getRange(destinyRange).setValues(data);
}


function launcher(){
  var link = 'https://docs.google.com/spreadsheets/d/1F57Ubru6sEwlj59keZ5RB_DaGVeevvn0dvQ016Lbu9E/edit#gid=1818396909'; //Link al collector
  var tabStaff = "Stats Staff"; // Pestaña de Staff en origen
  var tabUC = "Stats UC"; // Pestaña de UC en origen
  var tabAgStaff = "Aggregated Staff"; // Agregado de Staff en origen
  var tabAgUc = "Aggregated UC"; // Agregado de UC en origen

// Trae datos con columna de fecha
  importData(link,tabStaff,tabStaff);
  importData(link,tabUC,tabUC);
    
// Trae pestaña con fecha
  copyHereWithDate(link,tabAgStaff);
  copyHereWithDate(link,tabAgUc);
    
}