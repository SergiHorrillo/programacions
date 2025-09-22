/**
 * Flux integrat: Generació + Formatació en un sol arxiu.
 * Requisits: API Avançada de Google Docs activada (Servei "Google Docs API").
 * Funció principal a executar: generarIFormatar()
 * Aquest fitxer substitueix l'ús separat de Generarinforme.js i formatador.js.
 */

// ============================= FASE GENERACIÓ =============================

/**
 * Genera el document a partir del full "Dades" i les pestanyes Ln.
 * Retorna objecte amb docId i nomCopia.
 */
function generarDocumentPrograma() {
  var inici = Date.now();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fulla = spreadsheet.getSheetByName('Dades');
  if (!fulla) throw new Error('No s\'ha trobat la pestanya "Dades"');

  var dades = fulla.getRange('A1:S2').getValues();
  var valorE2 = fulla.getRange('E2').getValue();
  var valorF2 = fulla.getRange('F2').getValue();

  // Selecció plantilla
  var templateId = '';
  if (valorE2 == 1) {
    templateId = (valorF2 == 3) ? '15LHorRbkTuK0XGiRVFtLOJQYGCbpGqEBVsvc1Bur6o4' : '1hNc9qoNRvH8KKPHiT8BnEdwIVr-psDQyrzJLnftwzow';
  } else if (valorE2 == 2) {
    if (valorF2 == 4) templateId = '1tkUFPF7YgNiOtEUA9iGOEcZ9VDpLhsygtNwiDEPmlHs';
    else if (valorF2 == 5) templateId = '1-3RZzp8jS-CXOjzlfIVdm1vf9UdV_oWx24JEU_6tZz0';
  }
  if (!templateId) throw new Error('No hi ha plantilla per a E2=' + valorE2 + ' i F2=' + valorF2);

  var nomBase = fulla.getRange('B2').getValue();
  var ara = new Date();
  var dataFormatada = Utilities.formatDate(ara, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  var nomCopia = nomBase + ' (' + dataFormatada + ')';

  var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  var folderIterator = spreadsheetFile.getParents();
  var folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.getRootFolder();

  var copiaFile = DriveApp.getFileById(templateId).makeCopy(nomCopia, folder);
  var docId = copiaFile.getId();
  var document = DocumentApp.openById(docId);
  var cos = document.getBody();

  // Peu de pàgina
  var footer = document.getFooter();
  if (footer) footer.replaceText('<<NOM MATÈRIA>>', fulla.getRange('B2').getValue());

  // Variables generals
  var variables = {
    '<<NOM MATÈRIA>>': dades[1][1],
    '<<Departament>>': dades[1][0],
    '<<Tipus>>': dades[1][2],
    '<<Credits>>': dades[1][3],
    '<<Blocs>>': dades[1][4],
    '<<Lliuraments>>': dades[1][5]
  };
  for (var clau in variables) {
    cos.replaceText(clau, variables[clau]);
  }

  // Extres
  var extresSheet = spreadsheet.getSheetByName('Extres');
  if (extresSheet) {
    var lastRowExtres = extresSheet.getLastRow();
    var dataColA = extresSheet.getRange('A10:A' + lastRowExtres).getValues();
    var dataColB = extresSheet.getRange('B10:B' + lastRowExtres).getValues();
    var textEX = dataColA.map(function(r){return r[0];}).filter(Boolean).map(function(t){return '- ' + t;}).join('\n');
    var textAUT = dataColB.map(function(r){return r[0];}).filter(Boolean).map(function(t){return '- ' + t;}).join('\n');
    cos.replaceText('<<EX1L1B1>>', textEX);
    cos.replaceText('<<AUTL1B1>>', textAUT);
    cos.replaceText('<<EXCE>>', textEX);
    cos.replaceText('<<AUTORIA>>', textAUT);
  }

  // Taules per pestanyes Ln
  var numPestanyes = valorE2 * valorF2;
  for (var i = 1; i <= numPestanyes; i++) {
    var nomPestanya = 'L' + i;
    var sheetLn = spreadsheet.getSheetByName(nomPestanya);
    if (sheetLn) {
      var titol = sheetLn.getRange('B1:C1').getValue();
      cos.replaceText('<<Titol_' + nomPestanya + '>>', titol);
    }
    // Assumim que les funcions taulaA..D existeixen (estan definides en aquest mateix fitxer o un altre carregat)
    if (typeof taulaA === 'function') taulaA(nomPestanya, cos, spreadsheet);
    if (typeof taulaB === 'function') taulaB(nomPestanya, cos, spreadsheet);
    if (typeof taulaC === 'function') taulaC(nomPestanya, cos, spreadsheet);
    if (typeof taulaD === 'function') taulaD(nomPestanya, cos, spreadsheet);
  }

  // Placeholders PLn
  for (var p = 1; p <= numPestanyes; p++) {
    var valorPL = fulla.getRange(2, 8 + (p - 1)).getValue();
    cos.replaceText('<<PL' + p + '>>', valorPL);
  }

  // Imatges de portada
  if (typeof inserirImatgesPortada === 'function') {
    inserirImatgesPortada(cos, fulla);
  }

  document.saveAndClose();
  DocumentApp.flush();
  Logger.log('[GEN] Document generat en ' + (Date.now() - inici) + ' ms. DocID=' + docId);
  return { docId: docId, nomCopia: nomCopia };
}

// ============================= ESPERA DISPONIBILITAT =============================
/**
 * Espera fins que el document estigui accessible per l'API de Docs.
 * Retorna true si disponible o llança error després d'intents.
 */
function esperarDisponibilitatDoc(docId) {
  var intents = 7;
  for (var i = 0; i < intents; i++) {
    try {
      DocumentApp.openById(docId); // Si obre sense error, ja està
      if (i > 0) Logger.log('[WAIT] Document disponible al intent ' + (i + 1));
      return true;
    } catch (e) {
      Utilities.sleep(650 + (i * 120)); // backoff lleu
    }
  }
  throw new Error('Document no disponible després d\'esperar. DocID=' + docId);
}

// ============================= FORMATADOR (API DOCS) =============================
/**
 * Aplica amplades de columnes a taules A i D.
 */
function aplicarAmpladesTaules_(docId) {
  var puntsPerCm = 28.35;
  var docStruct = Docs.Documents.get(docId, { fields: 'body/content' });
  var allTables = docStruct.body.content.filter(function(e){ return !!e.table; });
  var widthRequests = [];
  var taulesA = 0, taulesD = 0;

  for (var i = 0; i < allTables.length; i++) {
    var tableElement = allTables[i];
    var table = tableElement.table;
    if (!table.tableRows || !table.tableRows[0] || !table.tableRows[0].tableCells) continue;
    var primerElement = (((table.tableRows[0]||{}).tableCells[0]||{}).content||[])[0];
    if (!primerElement || !primerElement.paragraph || !primerElement.paragraph.elements || !primerElement.paragraph.elements[0].textRun) continue;
    var firstCellContent = primerElement.paragraph.elements[0].textRun.content || '';
    var tableStart = tableElement.startIndex;

    if (firstCellContent.indexOf('BLOC') !== -1) { // Taula A
      taulesA++;
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [0], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 3 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [1], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 22 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [2], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 2.5 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
    } else if (firstCellContent.indexOf('Tipus d\'activitat') !== -1) { // Taula D
      taulesD++;
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [0], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 4 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [1], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 17.5 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [2], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 3 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
      widthRequests.push({ updateTableColumnProperties: { tableStartLocation: { index: tableStart }, columnIndices: [3], tableColumnProperties: { widthType: 'FIXED_WIDTH', width: { magnitude: 3 * puntsPerCm, unit: 'PT' } }, fields: 'width,widthType' } });
    }
  }

  if (widthRequests.length) {
    Docs.Documents.batchUpdate({ requests: widthRequests }, docId);
  }
  Logger.log('[FMT-WIDTH] Taules A: ' + taulesA + ' | Taules D: ' + taulesD + ' (aplicades amplades)');
}

/**
 * Combina celes de la columna ODS (files 1-3) en taules A.
 */
function combinarCelaTaulesA_(docId) {
  var docStruct = Docs.Documents.get(docId, { fields: 'body/content' });
  var allTables = docStruct.body.content.filter(function(e){ return !!e.table; });
  var mergeRequests = [];

  for (var i = 0; i < allTables.length; i++) {
    var tableElement = allTables[i];
    var table = tableElement.table;
    if (!table.tableRows || !table.tableRows[0] || !table.tableRows[0].tableCells) continue;
    var primerElement = (((table.tableRows[0]||{}).tableCells[0]||{}).content||[])[0];
    if (!primerElement || !primerElement.paragraph || !primerElement.paragraph.elements || !primerElement.paragraph.elements[0].textRun) continue;
    var firstCellContent = primerElement.paragraph.elements[0].textRun.content || '';

    if (firstCellContent.indexOf('BLOC') !== -1) {
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

  if (!mergeRequests.length) {
    Logger.log('[FMT-MERGE] Cap taula A trobada per combinar.');
    return;
  }
  // Ordenar invers per evitar invalidació d'índexs.
  mergeRequests.sort(function(a,b){
    return b.mergeTableCells.tableRange.tableCellLocation.tableStartLocation.index - a.mergeTableCells.tableRange.tableCellLocation.tableStartLocation.index;
  });

  Docs.Documents.batchUpdate({ requests: mergeRequests }, docId);
  Logger.log('[FMT-MERGE] Combinades ' + mergeRequests.length + ' taules A.');
}

/**
 * Aplica format complet (amplades + merges) amb tolerància a errors transitoris.
 */
function aplicarFormatador(docId) {
  Logger.log('[FMT] Inici formatador DocID=' + docId);
  // Fase amplades
  try {
    aplicarAmpladesTaules_(docId);
  } catch (e1) {
    Logger.log('[FMT-WIDTH][WARN] Primer intent ha fallat: ' + e1);
    Utilities.sleep(800);
    aplicarAmpladesTaules_(docId); // segon intent
  }
  // Fase merges
  try {
    combinarCelaTaulesA_(docId);
  } catch (e2) {
    Logger.log('[FMT-MERGE][WARN] Primer intent ha fallat: ' + e2);
    Utilities.sleep(800);
    combinarCelaTaulesA_(docId);
  }
  Logger.log('[FMT] Format complet aplicat.');
}

// ============================= ORQUESTADOR =============================
/**
 * Flux complet: generar -> esperar -> formatar.
 * Retorna el docId final.
 */
function generarIFormatar() {
  try {
    var resultat = generarDocumentPrograma();
    esperarDisponibilitatDoc(resultat.docId);
    aplicarFormatador(resultat.docId);
    Logger.log('[DONE] URL: https://docs.google.com/document/d/' + resultat.docId + '/edit');
    return resultat.docId;
  } catch (e) {
    Logger.log('[ERROR] ' + e);
    throw e;
  }
}

// ============================= COMPATIBILITAT ANTIGA =============================
/**
 * Wrapper legacy: manté el nom antic per scripts o triggers existents.
 * Recomanat usar generarIFormatar() a partir d'ara.
 */
function substituirValorsDocumentFinal() {
  return generarIFormatar();
}
