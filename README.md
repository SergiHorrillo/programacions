# Script integrat de generació i formatació de Programacions

Aquest repositori conté el flux unificat per generar un Document de Google (programació) a partir d'un full de càlcul font i aplicar-hi el format avançat (amplades de columnes i combinació de cel·les) en una sola execució.

## Arxiu principal

`generar_i_fomatar.js`  (nom històric mantingut) és l'únic script necessari actualment. Substitueix els antics `Generarinforme.js` i `formatador.js`.

Funció principal a executar: `generarIFormatar()`

També existeix un *wrapper* de compatibilitat: `substituirValorsDocumentFinal()` que simplement crida el flux nou per no trencar triggers o botons antics.

## Requisits previs

- Google Apps Script (entorn vinculat al full de càlcul de dades)
- Activar el servei avançat: **Google Docs API** (Editor > Serveis avançats de Google > habilitar) i també a la consola del projecte (necessari la primera vegada)
- Full de càlcul amb pestanya `Dades` i pestanyes de lliuraments `L1`, `L2`, ..., fins al màxim que correspongui.
- (Opcional) Pestanya `Extres` per omplir placeholders d'annexos.

## Estructura funcional

1. Selecció de plantilla: Depèn de les cel·les `E2` i `F2` de la pestanya `Dades` (multiplicitat de combinacions). Es centralitza a `CONFIG.selectTemplate(e2, f2)`.
2. Creació de còpia de plantilla dins la mateixa carpeta on resideix el full de càlcul.
3. Substitució de placeholders generals: `<<NOM MATÈRIA>>`, `<<Departament>>`, `<<Tipus>>`, etc.
4. Inserció de contingut per a cada pestanya `Ln`: construcció de taules A, B, C i D (funcions `taulaA..D`).
5. Inserció d'imatges ODS (icones) i imatges de portada si la funció `inserirImatgesPortada` està definida.
6. Tancament del document.
7. Espera controlada fins que el document sigui accessible per l'API avançada (evita condicions de carrera).
8. Format avançat amb l'API de Google Docs:
   - Ajust d'amplades de columnes a Taules A i D (unitats en punts derivades de cm via `CONFIG.PUNTS_PER_CM`).
   - Combinació vertical (merge) de la columna ODS (3 primeres files útils) a cada Taula A, amb lògica dinàmica si hi ha menys files.
9. Reintents: El formatador aplica fins a 3 intents amb *backoff* progressiu en cas d'errors transitoris.

## Objecte CONFIG

```js
var CONFIG = (function(){
  var PUNTS_PER_CM = 28.35; // conversió constant
  var TEMPLATES = { /* map de combinacions e2_f2 -> templateId */ };
  var ODS_MAP = { /* índex -> URL descarrega icona ODS */ };
  function selectTemplate(e2, f2){
    return TEMPLATES[e2 + '_' + f2] || TEMPLATES[e2 + '_default'] || null;
  }
  return { PUNTS_PER_CM: PUNTS_PER_CM, ODS_MAP: ODS_MAP, selectTemplate: selectTemplate };
})();
```

Integrar nous templates: afegir una clau al mapa `TEMPLATES` seguint el patró `E2_F2` o bé `E2_default` com a *fallback*.

## Funcions clau

- `generarDocumentPrograma()`: Construeix el document base (placeholders, taules, imatges) i retorna `{ docId, nomCopia }`.
- `esperarDisponibilitatDoc(docId)`: Intenta obrir via `DocumentApp` amb reintents; assegura que l'API avançada veurà totes les taules.
- `aplicarAmpladesTaules_(docId)`: Usa l'API Docs per establir amplades FIXED_WIDTH a Taules A i D.
- `combinarCelaTaulesA_(docId)`: Fusiona la columna ODS (files 1..n fins a 3) per a cada Taula A.
- `aplicarFormatador(docId)`: Orquestra amplades + merges amb fins a 3 intents.
- `generarIFormatar()`: Pipeline complet (crida recomanada).

## Placeholders principals

| Placeholder | Origen | Descripció |
|-------------|--------|-----------|
| `<<NOM MATÈRIA>>` | Dades!B2 | Nom de la matèria (també al peu) |
| `<<Departament>>` | Dades!A2 | Departament |
| `<<Tipus>>` | Dades!C2 | Tipus o modalitat |
| `<<Credits>>` | Dades!D2 | Crèdits |
| `<<Blocs>>` | Dades!E2 | Nombre de blocs (factor multiplicatiu) |
| `<<Lliuraments>>` | Dades!F2 | Nombre de lliuraments (factor multiplicatiu) |
| `<<PLn>>` | Dades!columnes H.. | Plaç/valor per lliurament n |

Els placeholders personalitzats de taules/imatges depenen de la implementació de `taulaA..D` i `inserirImatgesPortada`.

## Gestió d'errors i reintents

- El principal risc és intentar formatar abans que el document tingui totes les taules indexades. La combinació `esperarDisponibilitatDoc + pausa extra` minimitza el problema.
- Si un intent de format falla (ex.: 429, índex inconsistent), s'espera (backoff) i es reintenta fins a 3 vegades.
- El log mostra etiquetes: `[GEN]`, `[WAIT]`, `[FMT]`, `[FMT-WIDTH]`, `[FMT-MERGE]`, `[DONE]`, `[ERROR]`.

## Extensió futura (roadmap suggerit)

1. Lectures massives (batch) a `taulaB`, `taulaC`, `taulaA` per reduir crides `getRange` repetitives.
2. Caché de blobs d'icones ODS (evitar descarregar la mateixa URL múltiples vegades si es repeteix).
3. Validació de coherència: comprovar nombre esperat de taules = `numPestanyes * 4` abans de formatar (si es necessita més robustesa).
4. Marcatge idempotent (inserir un comentari invisible o marcador) per evitar format duplicat en execucions repetides.
5. Tests de regressió (unit test amb GAS *clasp* + entorn simulat) per funcions pures.

## Ús típic

1. Obrir el full de càlcul base.
2. Executar (menú Extensions > Apps Script) la funció `generarIFormatar`.
3. Revisar el nou document generat (es mostra URL al log `[DONE]`).

## Solució de problemes (FAQ)

| Símptoma | Possible causa | Acció |
|----------|----------------|-------|
| Error "No hi ha plantilla" | Combinació E2/F2 no mapejada | Afegir entrada a `TEMPLATES` o `*_default` |
| Amplades no aplicades | Document no llest | Augmentar pausa després d'`esperarDisponibilitatDoc` |
| Merge no fet en algunes taules | Taula amb <2 files de cos | Comportament esperat (no hi ha res a fusionar) |
| Execució lenta | Múltiples `getRange` dispersos | Implementar optimitzacions batch (roadmap) |
| Icònies ODS no apareixen | URL caducada o sense permís | Verificar que les URLs segueixen actives |

## Canvis respecte versió antiga

- Integració en un sol arxiu (abans: generació + formatador separats).
- Reintents automàtics per formatat.
- RowSpan dinàmic per a merges (abans fix a 3); ara s'adapta.
- Centralització de constants i plantilles amb `CONFIG`.
- Eliminació de `DocumentApp.flush()` (innecessari i problemàtic en alguns entorns).

## Llicència / Notes

Ús intern docent. Afegir llicència si es distribueix externament.

---
Qualsevol dubte o millora: crear una incidència o ampliar el roadmap.
