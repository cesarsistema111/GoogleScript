// @ts-nocheck
// https://spreadsheet.dev/objects-apps-script
// https://developers.google.com/apps-script/guides/macro-converter/convert-files
// https://www.youtube.com/watch?v=1qW5d7IAFBQ
// https://www.youtube.com/watch?v=J5fnTv-wnJw
function parametrosImport() {
  
  /** funcion donde se capturan los parametros de los nombres de la hojas */
  var importRanges = [
  {
    destinySheetName: "Destino", // hoja destino 
    fromFileKey: "1CBypMRn-wbnnPZ-86A2l_Kn9PzJ-xSU6LvUR3gi5Kbg", // id del hoja de calculo
    fromSheet:"BD",// hoja origen
    fromRange:"a:g" // rango de hoja origen

  },

  ];
 

importArray(importRanges)

}

function  importArray (importRanges) {
// funcion que sirve para recorrer el arreglo que llamamos imporRanges
// es  manera de ejemplo porque el arrglo solo tiene un elemento
for(var i=0; i< importRanges.length; i++){
    datExternaata(importRanges[i].fromFileKey,importRanges[i].destinySheetName,importRanges[i].fromSheet,importRanges[i].fromRange )

/*Logger.log( importRanges[i].fromSheet+' '+
            importRanges[i].fromRange+' '+
            importRanges[i].destinySheetName+' '+
            importRanges[i].fromFileKey);
*/

// mandemos estos datos a otra funcion para trabajarla

};
function datExternaata(idLibro,hojaDestino,hojaOrigen,rangoorigen ){

/**Seleccionar libro activo */
  var  libro=SpreadsheetApp.getActiveSpreadsheet();
/** nos situamos sobre la hoja de calculo "Destino" */
  var hoja01 =libro.getSheetByName(hojaDestino); 
  var ultimaFila=hoja01.getLastRow()+1;
  var data=SpreadsheetApp.openById(idLibro).getSheetByName(hojaOrigen).getRange(rangoorigen).getValues();
  Logger.log (ultimaFila)
  hoja01.getRange(ultimaFila,1,data.length,data[0].length).setValues(data);

}
}