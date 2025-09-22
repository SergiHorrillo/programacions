/**
 * Aquest script serveix per formatar les taules A i D d'un document de Google Docs JA EXISTENT.
 * S'executa manualment un cop el document ha estat creat.
 * Aquesta versió aplica primer les amplades i després combina les cel·les, ordenant les
 * peticions de manera inversa per garantir una execució sense errors.
 *
 * INSTRUCCIONS:
 * 1. Assegura't que l'API Avançada de Google Docs estigui activada a "Serveis".
 * 2. Canvia el 'docId' si vols formatar un altre document.
 * 3. Selecciona la funció "formatarTaulesDocumentExistent" al menú desplegable de l'editor.
 * 4. Fes clic a "Executa".
 */
function formatarTaulesDocumentExistent() {
  // IMPORTANT: Aquest és l'ID del document que vols formatar.
  var docId = "1PEfZZoV4xA2eK0p2clezue4dM5VmyyWHOtqxjwjy4AU";

  Logger.log("Iniciant formatació per al document ID: " + docId);
  
  try {
    // ===================================================================
    // FASE 1: APLICAR AMPLE DE COLUMNA (Aquesta part ja funcionava)
    // ===================================================================
    Logger.log("Fase 1: Aplicant amplada de columnes...");
    var docStruct = Docs.Documents.get(docId, {fields: 'body/content'});
    var allTables = docStruct.body.content.filter(function(e) { return e.table; });
    var widthRequests = [];
    var puntsPerCm = 28.35;
    var taulesA_Trobades = 0;
    var taulesD_Trobades = 0;

    for (var i = 0; i < allTables.length; i++) {
      var tableElement = allTables[i];
      var table = tableElement.table;
      
      if (table.tableRows && table.tableRows[0] && table.tableRows[0].tableCells && table.tableRows[0].tableCells[0] && table.tableRows[0].tableCells[0].content[0].paragraph.elements[0].textRun) {
        var firstCellContent = table.tableRows[0].tableCells[0].content[0].paragraph.elements[0].textRun.content;
        var tableStart = tableElement.startIndex;
        
        if (firstCellContent.includes("BLOC")) {
          taulesA_Trobades++;
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [0], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 3 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [1], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 22 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [2], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 2.5 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
        }
        else if (firstCellContent.includes("Tipus d'activitat")) {
          taulesD_Trobades++;
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [0], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 4 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [1], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 17.5 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [2], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 3 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
          widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [3], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 3 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
        }
      }
    }

    if (widthRequests.length > 0) {
      Docs.Documents.batchUpdate({ requests: widthRequests }, docId);
      Logger.log("Amplada de columna aplicada a " + taulesA_Trobades + " taules A i " + taulesD_Trobades + " taules D.");
    } else {
      Logger.log("No s'ha trobat cap taula per aplicar amplada.");
    }
    
    // ===================================================================
    // FASE 2: COMBINAR CEL·LES (Amb l'ordenació inversa per garantir l'èxit)
    // ===================================================================
    Logger.log("Fase 2: Preparant la combinació de cel·les...");
    docStruct = Docs.Documents.get(docId, {fields: 'body/content'});
    allTables = docStruct.body.content.filter(function(e) { return e.table; });
    var mergeRequests = [];
    
    for (var i = 0; i < allTables.length; i++) {
      var tableElement = allTables[i];
      var table = tableElement.table;
      
      if (table.tableRows && table.tableRows[0] && table.tableRows[0].tableCells && table.tableRows[0].tableCells[0] && table.tableRows[0].tableCells[0].content[0].paragraph.elements[0].textRun) {
        var firstCellContent = table.tableRows[0].tableCells[0].content[0].paragraph.elements[0].textRun.content;
        
        if (firstCellContent.includes("BLOC")) {
          mergeRequests.push({
            mergeTableCells: {
              tableRange: {
                tableCellLocation: {
                  tableStartLocation: { index: tableElement.startIndex },
                  rowIndex: 1,
                  columnIndex: 2
                },
                rowSpan: 3,
                columnSpan: 1
              }
            }
          });
        }
      }
    }

    if (mergeRequests.length > 0) {
      // LA CLAU: Ordenem les peticions de la darrera a la primera per no invalidar els índexs.
      mergeRequests.sort(function(a, b) {
        return b.mergeTableCells.tableRange.tableCellLocation.tableStartLocation.index - a.mergeTableCells.tableRange.tableCellLocation.tableStartLocation.index;
      });

      Docs.Documents.batchUpdate({ requests: mergeRequests }, docId);
      Logger.log("Combinació de cel·les aplicada a " + mergeRequests.length + " taules A.");
    } else {
      Logger.log("No s'han trobat taules A per combinar cel·les.");
    }
    
    Logger.log("Procés de formatació completat amb èxit.");

  } catch (e) {
    var missatgeExcepcio = "S'ha produït un error: " + e.toString();
    Logger.log(missatgeExcepcio);
  }
}

