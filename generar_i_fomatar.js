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

  // Selecció plantilla via CONFIG
  var templateId = CONFIG.selectTemplate(valorE2, valorF2);
  if (!templateId) throw new Error('No hi ha plantilla per a combinació E2=' + valorE2 + ' F2=' + valorF2);

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
  // Eliminat DocumentApp.flush(); en entorn on no està disponible i provocava TypeError
  Logger.log('[GEN] Document generat en ' + (Date.now() - inici) + ' ms. DocID=' + docId);
  return { docId: docId, nomCopia: nomCopia, numPestanyes: numPestanyes };
}

// ============================= CONFIG GLOBAL =============================
/**
 * Centralitza constants i selecció de plantilles / ODS.
 */
var CONFIG = (function(){
  var PUNTS_PER_CM = 28.35;
  var TEMPLATES = {
    '1_3': '15LHorRbkTuK0XGiRVFtLOJQYGCbpGqEBVsvc1Bur6o4',
    '1_default': '1hNc9qoNRvH8KKPHiT8BnEdwIVr-psDQyrzJLnftwzow',
    '2_4': '1tkUFPF7YgNiOtEUA9iGOEcZ9VDpLhsygtNwiDEPmlHs',
    '2_5': '1-3RZzp8jS-CXOjzlfIVdm1vf9UdV_oWx24JEU_6tZz0'
  };
  var ODS_MAP = { 1: "https://drive.google.com/uc?export=download&id=1HR8D87Kopm8hzarpICrylOKzX5AZJhG-", 2: "https://drive.google.com/uc?export=download&id=1rj4A7utzAxNgokWGgP5In_hAV0aWz16h", 3: "https://drive.google.com/uc?export=download&id=1WRBMYanemJm8QpOIrYdGE7CM4NfWo_Zt", 4: "https://drive.google.com/uc?export=download&id=116thhnZN-EftgAmk8epm1DCLysgB0NfR", 5: "https://drive.google.com/uc?export=download&id=1cnyGVYu_yiKVU-x-Z2rg6W_CH3b9giNB", 6: "https://drive.google.com/uc?export=download&id=1pBKpVC8BcplyQMSBK692Tj4eI4CwcOjC", 7: "https://drive.google.com/uc?export=download&id=1N8eCVXU7jDYrbBOJ-PLn4UTa4Txs1wq9", 8: "https://drive.google.com/uc?export=download&id=1MLQ_neg9vF0Dmn2IlS0YBUOcmF8Pd1_5", 9: "https://drive.google.com/uc?export=download&id=1TS5R6Gd8SXNKEC6JxNT4LdZds8GbbhNX", 10: "https://drive.google.com/uc?export=download&id=12FSqsOGriXFTiNAS4LOiYjAcRE1YZDmX", 11: "https://drive.google.com/uc?export=download&id=1jP0ON8z_u9h9XGNPyVDdlL20-ZORALgv", 12: "https://drive.google.com/uc?export=download&id=1yjiXhApkCm3VKu4FJV8JNmIyLN6fIdRb", 13: "https://drive.google.com/uc?export=download&id=1JGyCxDz9URl4TBDIicwYWiauhqo_ovYm", 14: "https://drive.google.com/uc?export=download&id=1Dn3-DWi9X73cGC4pAFyVTPt1OqHzG1Ao", 15: "https://drive.google.com/uc?export=download&id=1BBWN7-4y5XeDA0GSWv7GcCUKvm6AaA7o", 16: "https://drive.google.com/uc?export=download&id=11YIN8wNwlJNO4ltE7-W7bwV-K7cihG-f", 17: "https://drive.google.com/uc?export=download&id=1oYvcKtei-IgDRGe7EU_FMNCUb3wF8LvS", 18: "https://drive.google.com/uc?export=download&id=1ANKdpkSsl7mHv5dN1U_1N-gkUuGeoaFp" };
  function selectTemplate(e2, f2){
    return TEMPLATES[e2 + '_' + f2] || TEMPLATES[e2 + '_default'] || null;
  }
  return { PUNTS_PER_CM: PUNTS_PER_CM, ODS_MAP: ODS_MAP, selectTemplate: selectTemplate };
})();

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

// ============================= VERIFICACIÓ ESTRUCTURAL =============================
/**
 * Compta el nombre de taules presents al document (via Docs API).
 */
function comptarTaulesDoc_(docId) {
  var docStruct = Docs.Documents.get(docId, { fields: 'body/content' });
  if (!docStruct || !docStruct.body || !docStruct.body.content) return 0;
  var total = 0;
  for (var i = 0; i < docStruct.body.content.length; i++) {
    if (docStruct.body.content[i].table) total++;
  }
  return total;
}

/**
 * Espera fins que hi hagi com a mínim "esperades" taules (o llança error després d'intents).
 * Si el document té més taules (per plantilla), la condició es compleix igualment.
 */
function esperarEstructuraCompleta(docId, esperades) {
  var intents = 6;
  for (var i = 1; i <= intents; i++) {
    var compt = comptarTaulesDoc_(docId);
    if (compt >= esperades) {
      Logger.log('[WAIT-STRUCT] Taules detectades ' + compt + ' (>= ' + esperades + ') al intent ' + i);
      return compt;
    }
    var espera = 400 + (i - 1) * 300;
    Logger.log('[WAIT-STRUCT] Detectades ' + compt + '/' + esperades + ' taules. Esperant ' + espera + ' ms (intent ' + i + ')');
    Utilities.sleep(espera);
  }
  throw new Error('Estructura incompleta: no s\'han detectat ' + esperades + ' taules després d\'esperar.');
}

// ============================= FORMATADOR (API DOCS) =============================
/**
 * Aplica amplades de columnes a taules A i D.
 */
function aplicarAmpladesTaules_(docId) {
  var puntsPerCm = CONFIG.PUNTS_PER_CM;
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
      var totalRows = table.tableRows.length; // inclou capçalera
      var bodyRows = Math.max(0, totalRows - 1);
      var span = Math.min(3, bodyRows); // fins a 3 files o les disponibles
      if (span > 1) { // només té sentit merge si hi ha 2 o més files a combinar
        mergeRequests.push({
          mergeTableCells: {
            tableRange: {
              tableCellLocation: {
                tableStartLocation: { index: tableElement.startIndex },
                rowIndex: 1,
                columnIndex: 2
              },
              rowSpan: span,
              columnSpan: 1
            }
          }
        });
      }
    }
  }

  if (!mergeRequests.length) {
    Logger.log('[FMT-MERGE] Cap taula A per combinar (o massa curta).');
    return;
  }
  mergeRequests.sort(function(a,b){
    return b.mergeTableCells.tableRange.tableCellLocation.tableStartLocation.index - a.mergeTableCells.tableRange.tableCellLocation.tableStartLocation.index;
  });
  Docs.Documents.batchUpdate({ requests: mergeRequests }, docId);
  Logger.log('[FMT-MERGE] Combinades ' + mergeRequests.length + ' taules A (rowSpan dinàmic).');
}

/**
 * Aplica format complet (amplades + merges) amb tolerància a errors transitoris.
 */
function aplicarFormatador(docId) {
  Logger.log('[FMT] Inici formatador DocID=' + docId);
  var maxIntents = 3;
  for (var intent = 1; intent <= maxIntents; intent++) {
    var iniciIntent = Date.now();
    try {
      // Amplades
      try {
        aplicarAmpladesTaules_(docId);
      } catch (eWidth) {
        if (intent === maxIntents) throw new Error('Amplades fallides intent final: ' + eWidth);
        Logger.log('[FMT-WIDTH][WARN] Intent ' + intent + ' fallit: ' + eWidth);
        throw eWidth; // força pas al catch exterior per reintentar tot
      }
      // Merges
      try {
        combinarCelaTaulesA_(docId);
      } catch (eMerge) {
        if (intent === maxIntents) throw new Error('Merge fallit intent final: ' + eMerge);
        Logger.log('[FMT-MERGE][WARN] Intent ' + intent + ' fallit: ' + eMerge);
        throw eMerge;
      }
      var durada = Date.now() - iniciIntent;
      Logger.log('[FMT] Format aplicat correctament al intent ' + intent + ' (' + durada + ' ms)');
      return; // èxit -> sortim
    } catch (eTotal) {
      if (intent < maxIntents) {
        var espera = 900 + (intent - 1) * 700;
        Logger.log('[FMT][INFO] Reintent en ' + espera + ' ms...');
        Utilities.sleep(espera);
      } else {
        Logger.log('[FMT][ERROR] Fracàs després de ' + maxIntents + ' intents: ' + eTotal);
        throw eTotal;
      }
    }
  }
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
    // Verificació estructural: esperem el nombre mínim de taules (4 per pestanya)
    var esperades = resultat.numPestanyes * 4;
    try {
      esperarEstructuraCompleta(resultat.docId, esperades);
    } catch (eStruct) {
      // No aturem el procés, però avisem (pot ser que la plantilla tingui placeholders absents)
      Logger.log('[WARN][STRUCT] ' + eStruct);
    }
    Utilities.sleep(300); // breu pausa final abans del formatador
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
