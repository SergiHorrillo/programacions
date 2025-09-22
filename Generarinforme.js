/**
 * Script original per generar la programació, amb la nova funcionalitat
 * per inserir les imatges de portada i avaluació.
 * CORREGIT: Mètode d'extracció d'ID d'URL millorat i més robust.
 */

// ===================================================================
// FUNCIÓ PRINCIPAL (Sense canvis)
// ===================================================================
function substituirValorsDocumentFinal() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fulla = spreadsheet.getSheetByName("Dades");
  var dades = fulla.getRange("A1:S2").getValues(); // Ampliem el rang per incloure R2 i S2

  // Llegir E2 i F2 d'una sola vegada
  var valorE2 = fulla.getRange("E2").getValue();
  var valorF2 = fulla.getRange("F2").getValue();

  // Seleccionar la plantilla segons E2 i, si cal, F2
  var templateId = "";
  if (valorE2 == 1) {
    templateId = (valorF2 == 3) ? "15LHorRbkTuK0XGiRVFtLOJQYGCbpGqEBVsvc1Bur6o4" : "1hNc9qoNRvH8KKPHiT8BnEdwIVr-psDQyrzJLnftwzow";
  } else if (valorE2 == 2) {
    if (valorF2 == 4) templateId = "1tkUFPF7YgNiOtEUA9iGOEcZ9VDpLhsygtNwiDEPmlHs";
    else if (valorF2 == 5) templateId = "1-3RZzp8jS-CXOjzlfIVdm1vf9UdV_oWx24JEU_6tZz0";
  }
  if (!templateId) {
    throw new Error("No hi ha plantilla per a E2=" + valorE2 + " i F2=" + valorF2);
  }

  // Afegir data i hora al nom
  var nomBase = fulla.getRange("B2").getValue();
  var ara = new Date();
  var dataFormatada = Utilities.formatDate(ara, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  var nomCopia = nomBase + " (" + dataFormatada + ")";

  var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  var folderIterator = spreadsheetFile.getParents();
  var folder = (folderIterator.hasNext()) ? folderIterator.next() : DriveApp.getRootFolder();

  var plantilla = DriveApp.getFileById(templateId);
  var copiaFile = plantilla.makeCopy(nomCopia, folder);

  var document = DocumentApp.openById(copiaFile.getId());
  var cos = document.getBody();

  // Substituir el placeholder del peu de pàgina amb el valor de la cel·la B2 de "Dades"
  var footer = document.getFooter();
  if (footer) {
    footer.replaceText("<<NOM MATÈRIA>>", fulla.getRange("B2").getValue());
  }

  // Substituir etiquetes generals
  var variables = { '<<NOM MATÈRIA>>': dades[1][1], '<<Departament>>': dades[1][0], '<<Tipus>>': dades[1][2], '<<Credits>>': dades[1][3], '<<Blocs>>': dades[1][4], '<<Lliuraments>>': dades[1][5] };
  for (var clau in variables) {
    cos.replaceText(clau, variables[clau]);
  }

  // Processar els placeholders d'"Extres"
  var extresSheet = spreadsheet.getSheetByName("Extres");
  var lastRowExtres = extresSheet.getLastRow();
  var dataColA = extresSheet.getRange("A10:A" + lastRowExtres).getValues();
  var dataColB = extresSheet.getRange("B10:B" + lastRowExtres).getValues();

  var textEX = dataColA.map(function(row) { return row[0]; }).filter(Boolean).map(function(item) { return "- " + item; }).join("\n");
  var textAUT = dataColB.map(function(row) { return row[0]; }).filter(Boolean).map(function(item) { return "- " + item; }).join("\n");
  
  cos.replaceText("<<EX1L1B1>>", textEX);
  cos.replaceText("<<AUTL1B1>>", textAUT);
  cos.replaceText("<<EXCE>>", textEX);
  cos.replaceText("<<AUTORIA>>", textAUT);

  // Bucle per iterar sobre pestanyes
  var numPestanyes = valorE2 * valorF2;
  for (var i = 1; i <= numPestanyes; i++) {
    var nomPestanya = "L" + i;
    var sheetLn = spreadsheet.getSheetByName(nomPestanya);
    if (sheetLn) {
      var titol = sheetLn.getRange("B1:C1").getValue();
      var placeholderTitol = "<<Titol_" + nomPestanya + ">>";
      cos.replaceText(placeholderTitol, titol);
    }
    
    taulaA(nomPestanya, cos, spreadsheet);
    taulaB(nomPestanya, cos, spreadsheet);
    taulaC(nomPestanya, cos, spreadsheet);
    taulaD(nomPestanya, cos, spreadsheet);
  }

  // Nou bloc: Substituir els placeholders <<PLn>>
  for (var i = 1; i <= numPestanyes; i++) {
    var valorPL = fulla.getRange(2, 8 + (i - 1)).getValue();
    var placeholderPL = "<<PL" + i + ">>";
    cos.replaceText(placeholderPL, valorPL);
  }

  inserirImatgesPortada(cos, fulla);

  document.saveAndClose();
  Logger.log("Document actualitzat: " + document.getUrl());
}


// ===================================================================
// == NOVA FUNCIÓ PER INSERIR IMATGES (AMB CORRECCIÓ) ==
// ===================================================================
/**
 * Insereix les imatges IMG1 i IMG2 al document.
 * @param {Body} cos El cos del document de Google Docs.
 * @param {Sheet} fulla La pestanya "Dades" del full de càlcul.
 */
function inserirImatgesPortada(cos, fulla) {
  /**
   * Funció auxiliar per extreure l'ID d'un URL de Google Drive.
   * Aquesta versió és més robusta i fiable.
   */
  function extreureIdDeUrl(url) {
    if (!url || typeof url !== 'string') return null;
    try {
      // Tallem l'URL per les barres i busquem el segment que correspon a l'ID
      return url.split('/d/')[1].split('/')[0];
    } catch (e) {
      Logger.log("No s'ha pogut extreure l'ID de l'URL: " + url);
      return null;
    }
  }

  var puntsPerCm = 28.35;

  // --- Processar la imatge de portada (IMG1) ---
  var urlImg1 = fulla.getRange("R2").getValue();
  var idImg1 = extreureIdDeUrl(urlImg1);
  
  if (idImg1) {
    var placeholder1 = cos.findText("<<IMG1>>");
    if (placeholder1) {
      var element1 = placeholder1.getElement();
      var parent1 = element1.getParent();
      parent1.clear(); // Buidem el paràgraf per deixar espai a la imatge
      
      try {
        var imatge1 = DriveApp.getFileById(idImg1).getBlob();
        var imatgeInserida1 = parent1.appendInlineImage(imatge1);
        
        // Ajustem la mida a 9x9 cm
        imatgeInserida1.setWidth(9 * puntsPerCm);
        imatgeInserida1.setHeight(9 * puntsPerCm);
      } catch (e) {
        Logger.log("Error en inserir la imatge IMG1 (ID: " + idImg1 + "). Assegura't que l'arxiu existeix i tens permisos. Error: " + e.message);
        parent1.appendText("<<ERROR: No s'ha pogut carregar la imatge IMG1>>");
      }
    }
  }

  // --- Processar la imatge d'avaluació (IMG2) ---
  var urlImg2 = fulla.getRange("S2").getValue();
  var idImg2 = extreureIdDeUrl(urlImg2);
  
  if (idImg2) {
    var placeholder2 = cos.findText("<<IMG2>>");
    if (placeholder2) {
      var element2 = placeholder2.getElement();
      var parentCell = element2.getParent().getParent().asTableCell(); // El pare és el paràgraf, l'avi és la cel·la
      parentCell.clear();
      
      try {
        var imatge2 = DriveApp.getFileById(idImg2).getBlob();
        var imatgeInserida2 = parentCell.insertImage(0, imatge2);
        
        // Ajustem a un ample de 15 cm mantenint la proporció
        var ampleFix2 = 15 * puntsPerCm;
        var ratio2 = imatgeInserida2.getHeight() / imatgeInserida2.getWidth();
        imatgeInserida2.setWidth(ampleFix2);
        imatgeInserida2.setHeight(ampleFix2 * ratio2);
      } catch (e) {
          Logger.log("Error en inserir la imatge IMG2 (ID: " + idImg2 + "). Assegura't que l'arxiu existeix i tens permisos. Error: " + e.message);
          parentCell.setText("<<ERROR: No s'ha pogut carregar la imatge IMG2>>");
      }
    }
  }
}


// ===================================================================
// == RESTA DE FUNCIONS (SENSE CANVIS) ==
// ===================================================================

function formatarCapcalera(taula, colorFons) { /* ... (codi original) ... */ }
function eliminarPlaceholder(cos, placeholder) { /* ... (codi original) ... */ }
function formatejarTextParcial(cell, start, end, fontSize) { /* ... (codi original) ... */ }
function taulaA(nomPestanya, cos, spreadsheet) { /* ... (codi original) ... */ }
function taulaB(nomPestanya, cos, spreadsheet) { /* ... (codi original) ... */ }
function taulaC(nomPestanya, cos, spreadsheet) { /* ... (codi original) ... */ }
function taulaD(nomPestanya, cos, spreadsheet) { /* ... (codi original) ... */ }

// (Aquí aniria el codi complet de les funcions originals que no es modifiquen)
// Per brevetat, no les reprodueixo totes, però al teu arxiu hi haurien de ser.
// Assegura't de copiar la funció `inserirImatgesPortada` i la crida a `substituirValorsDocumentFinal`.

// (Aquí enganxo les funcions completes per si de cas)

function formatarCapcalera(taula, colorFons) {
  var capcalera = taula.getRow(0);
  for (var c = 0; c < capcalera.getNumCells(); c++) {
    var cell = capcalera.getCell(c);
    cell.setBackgroundColor(colorFons);
    var childCount = cell.getNumChildren();
    for (var k = 0; k < childCount; k++) {
      var child = cell.getChild(k);
      if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        var paragraph = child.asParagraph();
        paragraph.setBold(true);
        paragraph.setFontSize(12);
      }
    }
  }
}

function eliminarPlaceholder(cos, placeholder) {
  var found = cos.findText(placeholder);
  if (!found) return null;
  var element = found.getElement().getParent();
  var index = cos.getChildIndex(element);
  element.removeFromParent();
  return index;
}

function formatejarTextParcial(cell, start, end, fontSize) {
  var childCount = cell.getNumChildren();
  if (childCount > 0) {
    for (var i = 0; i < childCount; i++) {
      var child = cell.getChild(i);
      if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        var text = child.asParagraph().editAsText();
        if (text.getText().length >= end) {
          text.setBold(start, end - 1, true);
          if (fontSize) text.setFontSize(start, end - 1, fontSize);
        }
        break;
      }
    }
  }
}

function taulaA(nomPestanya, cos, spreadsheet) {
  var index = eliminarPlaceholder(cos, "<<PRO_" + nomPestanya + "_A>>");
  if (index === null) return;

  var fullaSheet = spreadsheet.getSheetByName(nomPestanya);
  var fullaDades = spreadsheet.getSheetByName("Dades");

  var valorC3 = parseFloat(fullaSheet.getRange("C3").getValue());
  var valorF5 = parseFloat(fullaDades.getRange("F2").getValue());
  var textBloc = (valorC3 > valorF5) ? "BLOC II" : "BLOC I";
  var colorFons = (textBloc === "BLOC I") ? "#FFDDDD" : "#DDDDFF";

  var taulaContingut = [
    [textBloc, "LLIURAMENT " + valorC3 + ": " + fullaSheet.getRange("B1").getValue(), "ODS"],
    ["CONTEXT:", fullaSheet.getRange("A3").getValue(), ""],
    ["REPTE:", fullaSheet.getRange("B3").getValue(), ""],
    ["Components Competencials:", "", ""]
  ];

  var textCompetencials = "";
  for (var fila = 20; fila <= 22; fila++) {
    var valorD = fullaSheet.getRange(fila, 4).getValue();
    if (valorD) textCompetencials += "- " + valorD + "\n";
  }
  taulaContingut[3][1] = textCompetencials.trim();

  var taula = cos.insertTable(index, taulaContingut);
  formatarCapcalera(taula, colorFons);

  // Normalitzar mida de lletra del cos (evitar herències de 3 pt). No toquem la fila 0 (capçalera)
  for (var rNorm = 1; rNorm < taula.getNumRows(); rNorm++) {
    var filaNorm = taula.getRow(rNorm);
    for (var cNorm = 0; cNorm < filaNorm.getNumCells(); cNorm++) {
      var cellNorm = filaNorm.getCell(cNorm);
      for (var ch = 0; ch < cellNorm.getNumChildren(); ch++) {
        var fill = cellNorm.getChild(ch);
        if (fill.getType() === DocumentApp.ElementType.PARAGRAPH) {
          fill.asParagraph().setFontSize(11); // Assumpció: mida desitjada 11 pt
        }
      }
    }
  }

  for (var r = 1; r < taula.getNumRows(); r++) {
    taula.getRow(r).getCell(0).getChild(0).asParagraph().setBold(true);
  }

  // Eliminar negrita de la segona columna (índex 1) en les files del cos (r >= 1)
  for (var r2 = 1; r2 < taula.getNumRows(); r2++) {
    var cell2 = taula.getRow(r2).getCell(1);
    for (var k2 = 0; k2 < cell2.getNumChildren(); k2++) {
      var child2 = cell2.getChild(k2);
      if (child2.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var text2 = child2.asParagraph().editAsText();
        var len2 = text2.getText().length;
        if (len2 > 0) {
          text2.setBold(0, len2 - 1, false);
        }
      }
    }
  }

  var odsMap = { 1: "https://drive.google.com/uc?export=download&id=1HR8D87Kopm8hzarpICrylOKzX5AZJhG-", 2: "https://drive.google.com/uc?export=download&id=1rj4A7utzAxNgokWGgP5In_hAV0aWz16h", 3: "https://drive.google.com/uc?export=download&id=1WRBMYanemJm8QpOIrYdGE7CM4NfWo_Zt", 4: "https://drive.google.com/uc?export=download&id=116thhnZN-EftgAmk8epm1DCLysgB0NfR", 5: "https://drive.google.com/uc?export=download&id=1cnyGVYu_yiKVU-x-Z2rg6W_CH3b9giNB", 6: "https://drive.google.com/uc?export=download&id=1pBKpVC8BcplyQMSBK692Tj4eI4CwcOjC", 7: "https://drive.google.com/uc?export=download&id=1N8eCVXU7jDYrbBOJ-PLn4UTa4Txs1wq9", 8: "https://drive.google.com/uc?export=download&id=1MLQ_neg9vF0Dmn2IlS0YBUOcmF8Pd1_5", 9: "https://drive.google.com/uc?export=download&id=1TS5R6Gd8SXNKEC6JxNT4LdZds8GbbhNX", 10: "https://drive.google.com/uc?export=download&id=12FSqsOGriXFTiNAS4LOiYjAcRE1YZDmX", 11: "https://drive.google.com/uc?export=download&id=1jP0ON8z_u9h9XGNPyVDdlL20-ZORALgv", 12: "https://drive.google.com/uc?export=download&id=1yjiXhApkCm3VKu4FJV8JNmIyLN6fIdRb", 13: "https://drive.google.com/uc?export=download&id=1JGyCxDz9URl4TBDIicwYWiauhqo_ovYm", 14: "https://drive.google.com/uc?export=download&id=1Dn3-DWi9X73cGC4pAFyVTPt1OqHzG1Ao", 15: "https://drive.google.com/uc?export=download&id=1BBWN7-4y5XeDA0GSWv7GcCUKvm6AaA7o", 16: "https://drive.google.com/uc?export=download&id=11YIN8wNwlJNO4ltE7-W7bwV-K7cihG-f", 17: "https://drive.google.com/uc?export=download&id=1oYvcKtei-IgDRGe7EU_FMNCUb3wF8LvS", 18: "https://drive.google.com/uc?export=download&id=1ANKdpkSsl7mHv5dN1U_1N-gkUuGeoaFp" };

  for (var i = 0; i < 3; i++) {
    var valorE = fullaSheet.getRange(20 + i, 5).getValue();
    if (!valorE || String(valorE).trim() === ".") continue;
    var odsNum = parseInt(String(valorE).split('.')[0].trim());
    if (!isNaN(odsNum) && odsMap[odsNum]) {
      var celODS = taula.getRow(i + 1).getCell(2);
      celODS.clear();
      var blobImatge = UrlFetchApp.fetch(odsMap[odsNum]).getBlob();
      celODS.insertImage(0, blobImatge);
    }
  }
}

function taulaB(nomPestanya, cos, spreadsheet) {
  var index = eliminarPlaceholder(cos, "<<PRO_" + nomPestanya + "_B>>");
  if (index === null) return;
  var fullaSheet = spreadsheet.getSheetByName(nomPestanya);
  var taulaContingut = [["Objectius d'Aprenentatge", "CE - Criteris d'Avaluació"]];
  for (var fila = 6; fila <= 15; fila++) {
    var valorB = fullaSheet.getRange(fila, 2).getValue();
    if (valorB) {
      var valorC = fullaSheet.getRange(fila, 3).getValue();
      var valorD = fullaSheet.getRange(fila, 4).getValue();
      taulaContingut.push([valorB, valorC + " - " + valorD]);
    }
  }
  var taula = cos.insertTable(index, taulaContingut);
  formatarCapcalera(taula, "#DDDDDD");
  for (var r = 1; r < taula.getNumRows(); r++) {
    var cell = taula.getRow(r).getCell(1);
    var textTotal = cell.getText();
    var indexGuio = textTotal.indexOf(" - ");
    if (indexGuio > 0) {
      formatejarTextParcial(cell, 0, indexGuio, 14);
    }
  }
}

function taulaC(nomPestanya, cos, spreadsheet) {
  var index = eliminarPlaceholder(cos, "<<PRO_" + nomPestanya + "_C>>");
  if (index === null) return;
  var fullaSheet = spreadsheet.getSheetByName(nomPestanya);
  var sabersText = "";
  for (var fila = 20; fila <= 29; fila++) {
    var valorB = fullaSheet.getRange(fila, 2).getValue();
    if (valorB) sabersText += "\n- " + valorB;
  }
  sabersText = sabersText.trim();
  var taula = cos.insertTable(index, [["Sabers"], [sabersText]]);
  formatarCapcalera(taula, "#DDDDDD");
}

function taulaD(nomPestanya, cos, spreadsheet) {
  var index = eliminarPlaceholder(cos, "<<PRO_" + nomPestanya + "_D>>");
  if (index === null) return;
  var fullaSheet = spreadsheet.getSheetByName(nomPestanya);
  var dades = fullaSheet.getRange("D32:G59").getValues();
  var taulaContingut = [["Tipus d'activitat", "Activitat", "Aval. Sumativa(%)", "Aval. Formadora"]];
  for (var i = 0; i < dades.length; i++) {
    var tipusActivitat = dades[i][0];
    if (tipusActivitat) {
      var avalFormadora = dades[i][3] ? "SÍ" : "NO";
      taulaContingut.push([tipusActivitat, dades[i][1], dades[i][2], avalFormadora]);
    }
  }
  var taula = cos.insertTable(index, taulaContingut);
  formatarCapcalera(taula, "#DDDDDD");
  
  for (var r = 1; r < taula.getNumRows(); r++) {
    taula.getRow(r).getCell(0).getChild(0).asParagraph().setBold(true);
  }
  
  var parent = cos;
  var tableIndex = parent.getChildIndex(taula);
  parent.insertPageBreak(tableIndex + 1);
}
