/**
 * Aquest script prova la inserció de les dues imatges, <<IMG1>> i <<IMG2>>,
 * amb les seves respectives mides i configuracions.
 */
function provarInserirImatge() {
  // ===================================================================
  // PAS 1: Enganxa aquí l'ID del teu document de prova
  // ===================================================================
  var docIdDeProva = "1GD44EXcVN-84GML0KyDtVgFTHofPROHbtVlCo1e7ROM";
  // ===================================================================

  // URLs de les dues imatges
  var imageUrl1 = "https://drive.google.com/file/d/1mFwwCjKnyUf2H8U39AaLTgHA8RVfRZtP/view?usp=sharing";
  var imageUrl2 = "https://drive.google.com/file/d/1D1HXsrqfaJHD4tpJ40fvZkg7qQb7tB7w/view?usp=drive_link";

  if (docIdDeProva === "ID_DEL_TEU_DOCUMENT_DE_PROVA" || docIdDeProva === "") {
    SpreadsheetApp.getUi().alert("ERROR: Si us plau, edita el codi i substitueix 'ID_DEL_TEU_DOCUMENT_DE_PROVA' per l'ID real del teu document de prova.");
    return;
  }

  Logger.log("Iniciant prova d'inserció de dues imatges...");
  
  try {
    var document = DocumentApp.openById(docIdDeProva);
    var cos = document.getBody();
    var puntsPerCm = 28.35;
    
    function extreureIdDeUrl(url) {
      if (!url || typeof url !== 'string') return null;
      try {
        return url.split('/d/')[1].split('/')[0];
      } catch (e) { return null; }
    }

    // ===================================================================
    // Processar la imatge de portada (<<IMG1>>)
    // ===================================================================
    var idImg1 = extreureIdDeUrl(imageUrl1);
    Logger.log("ID extret per IMG1: " + idImg1);
    if (idImg1) {
      var placeholder1 = cos.findText("<<IMG1>>");
      if (placeholder1) {
        Logger.log("Placeholder '<<IMG1>>' trobat.");
        var elementText1 = placeholder1.getElement().asText();
        var parentParagraph1 = elementText1.getParent().asParagraph();
        
        var imatge1 = DriveApp.getFileById(idImg1).getBlob();
        var imatgeInserida1 = parentParagraph1.insertInlineImage(0, imatge1);
        elementText1.deleteText(0, elementText1.getText().length - 1);

        // Mida fixa de 14x14 cm
        imatgeInserida1.setWidth(14 * puntsPerCm);
        imatgeInserida1.setHeight(14 * puntsPerCm);
        Logger.log("Imatge IMG1 inserida i ajustada a 14x14 cm.");
      } else {
        Logger.log("Avís: No s'ha trobat el placeholder '<<IMG1>>'.");
      }
    }

    // ===================================================================
    // Processar la imatge d'avaluació (<<IMG2>>)
    // ===================================================================
    var idImg2 = extreureIdDeUrl(imageUrl2);
    Logger.log("ID extret per IMG2: " + idImg2);
    if (idImg2) {
      var placeholder2 = cos.findText("<<IMG2>>");
      if (placeholder2) {
        Logger.log("Placeholder '<<IMG2>>' trobat.");
        var elementText2 = placeholder2.getElement().asText();
        var parentParagraph2 = elementText2.getParent().asParagraph();

        var imatge2 = DriveApp.getFileById(idImg2).getBlob();
        var imatgeInserida2 = parentParagraph2.insertInlineImage(0, imatge2);
        elementText2.deleteText(0, elementText2.getText().length - 1);
        
        var parentCell2 = parentParagraph2.getParent().asTableCell();
        
        // ===================================================================
        // CANVI: Tornem a l'amplada de 16 cm
        // ===================================================================
        var ampladaDesitjadaCm = 16; // <-- MIDA REDUÏDA
        // ===================================================================
        
        var ampleFinalEnPunts = ampladaDesitjadaCm * puntsPerCm;
        var ratio = imatgeInserida2.getHeight() / imatgeInserida2.getWidth();
        imatgeInserida2.setWidth(ampleFinalEnPunts);
        imatgeInserida2.setHeight(ampleFinalEnPunts * ratio);
        Logger.log("Imatge IMG2 inserida i ajustada a " + ampladaDesitjadaCm + " cm d'ample.");

      } else {
        Logger.log("Avís: No s'ha trobat el placeholder '<<IMG2>>'.");
      }
    }
    
    Logger.log(">>> PROVA FINALITZADA! <<<");

  } catch (e) {
      Logger.log(">>> S'HA PRODUÏT UN ERROR FATAL: " + e.toString() + ".");
  }
}