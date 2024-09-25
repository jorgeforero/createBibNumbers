/**
 * CreateBibs
 * Generación de números de competencia
 */

// ID del Template de Slides 
const BIBTPLID = 'XX_INCLUYA_EL_ID_TEMPLATE_XX';
// ID del Folder para almacenar los Bibs generados
const BIBFLID = 'XX_ID_FOLDER_XX';

/**
 * createBibNumbers
 * Crea los números de competencia de los participantes de la carrera a partir de la información registrada
 * en la hoja de cálculo.  Genera archivos PDF con cada uno de los números generados
 * 
 * @param {void} - void
 * @return {void} - Bib numbers generados de acuerdo a los datos de los corredores
 */
function createBibNumbers() {
  try {
    let counter = 0;
    // Carga la información de los participantes desde la hoja de Cálculo
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'Runners' );    
    let runners = sheet.getDataRange().getDisplayValues();
    let header = runners.shift();
    // Recorre las información de los corredores para generar los números de completencia
    for ( let rndx=0; rndx<runners.length; rndx++ ) {
      let competitor = getRowAsObject( runners[ rndx], header );
      // Valida que se le haya asignado un número para poder general el Bib Number
      if ( competitor.bibnumber !== '' ) {
        // Generación del BIB         
        getBibNumberPDF_( BIBFLID, competitor.token, competitor.name.toUpperCase(), competitor.distance, competitor.bibnumber );
        counter++;
      };//else no se evalua - No se le ha asigando número
    };//for
    console.log( `Total Bibs Generados ${counter}` );
  } catch ( e ) {
    // Mensaje de error
    console.log( `Error: ${e.message}` );
  };
};

/**
 * getRowAsObject
 * Obtiene un objeto con los valores de la fila dada: RowData. Toma los nombres de las llaves del parámtero Header. Las llaves
 * son dadas en minusculas y los espacios reemplazados por _
 * 
 * @param {array} RowData - Arreglo con los datos de la fila de la hoja
 * @param {array} Header - Arreglo con los nombres del encabezado de la hoja
 * @return {object} obj - Objeto con los datos de la fila y las propiedades nombradas de acuerdo a Header
 */
 function getRowAsObject( RowData, Header ) {
  let obj = {};
  for ( let indx=0; indx<RowData.length; indx++ ) {
    obj[ Header[ indx ].toLowerCase().replace( /\s/g, '_' ) ] = RowData[ indx ];
  };//for
  return obj;
};

/**
 * getBibNumberPDF
 * Genera un archivo PDF a partir del template ( Google Slides ) y los datos del corredor. El archivo
 * generado es guardado en Drive ( Folder dado )
 * 
 * @param {string} FolderId - Id del folder donde se guardan los Bib numbers generados
 * @param {string} Token - Identificador del corredor
 * @param {string} Name - Nombre del corredor
 * @param {string} Distance - Distancia a la que está inscrito el corredor
 * @param {string} BibNumber - Númedor asignado al corredor
 * @return {void} - Archivo PDF generado de acuerdo a los datos de entrada
 */
function getBibNumberPDF_( FolderId, Token, Name, Distance, BibNumber ) {
  try {
    // Genera archivo nuevo a partir del template
    let templateFile = DriveApp.getFileById( BIBTPLID );
    let newFileId = templateFile.makeCopy().getId();
    let newDeck = SlidesApp.openById( newFileId );
    // Obtiene el slide template
    let newSlide = newDeck.getSlides()[ 0 ];
    // Obtiene las shapes del template
    let shapes = newSlide.getShapes();
    // Reemplazo de los placeholder con los datos dados
    shapes[ 0 ].getText().replaceAllText( '{{NUM}}', BibNumber );
    shapes[ 1 ].getText().replaceAllText( '{{NAME}}', Name );
    shapes[ 2 ].getText().replaceAllText( '{{DIST}}', Distance );
    // Salva los cambios en el nuevo slide
    newDeck.saveAndClose();
    // Url para obtener el pdf del certificado generado
    let url_base = 'https://docs.google.com/presentation/d/' + newFileId + '/export/pdf';
    // Opciones para el llamado del slide
    let options = {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    // Obtiene el json del slide generado
    let response = UrlFetchApp.fetch( url_base, options );
    let pdfName = Token + '_BN_1RACARRERA' + '.pdf';
    let blob = response.getBlob();
    blob.setName( pdfName );
    // Genera el PDF en el folder dado
    let folder = DriveApp.getFolderById( FolderId );
    let bibFile = folder.createFile( blob );
    // Permisos de acceso a internet
    bibFile.setSharing( DriveApp.Access.ANYONE, DriveApp.Permission.VIEW );
    // Remueve el archivo Slides recibido por parametro - Borrado directo por el API
    Drive.Files.remove( newFileId );
    // Resultado
    console.log( `BibNumber Generado! idfile= ${bibFile.getId()}`);
  } catch ( e ) {
    // Mensaje de error
    console.log( `Error: ${e.message}` );
  };
};
