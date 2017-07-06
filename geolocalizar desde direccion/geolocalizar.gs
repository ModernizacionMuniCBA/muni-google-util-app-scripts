var ui = SpreadsheetApp.getUi();

// columna con los datos a geolocalizar (se espera algo como "calle numero, ciudad, provincia, pais")
var addressColumn = 1;

// definicion de las columnas que se van a completar
var latColumn = 2;
var lngColumn = 3;
var foundAddressColumn = 4;
var qualityColumn = 5;
var sourceColumn = 6;

// definir la región a usar en la geolocalización
googleGeocoder = Maps.newGeocoder().setRegion(
  PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'ar'
);

// funciion para comenzar a Geolocalizar una lista de direcciones seleccionadas
function geocode(source) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();

  // es requisito que estén seleccionadas 6 columnas (sin importar el número de filas)
  if (cells.getNumColumns() != 6) {
    ui.alert(
      'Warning',
      'Tu selección debe contener 6 columnas',
      ui.ButtonSet.OK
    );
    return;
  }

  var nAll = 0;  // contador total
  var nIgnores = 0;  // contador de ignorados (ya tenian lat y long)
  var nFailure = 0;  // contador de fallados
  var quality;
  var printComplete = true;

  // iterar por todas las filas elegidas (cada una con una direccion a geolocalizar)
  for (addressRow = 1; addressRow <= cells.getNumRows(); addressRow++) {
    var address = cells.getCell(addressRow, addressColumn).getValue();

    nAll++;
    
    if (!address) {
      nIgnores++;
      continue;
    }
    
    // ignorar los que ya están
    var lat = cells.getCell(addressRow, latColumn).getValue();
    if (lat !== '') {
      nIgnores++;
      continue;
    }
    
    if (source == 'Google') {
      nFailure += withGoogle(cells, addressRow, address);
      // esperar un poco más de un segundo (revisar como afecta al límite)
      Utilities.sleep(1100);
    }
  }

  if (printComplete) {
    ui.alert('Completado!', 'Geocodificados: ' + (nAll - nFailure)
    + '\nFallados: ' + nFailure + ' Ignorados: ' + nIgnores, ui.ButtonSet.OK);
  }

}

/**
 * Geocode address with Google Apps https://developers.google.com/apps-script/reference/maps/geocoder
 */
function withGoogle(cells, row, address) {
  Logger.log('Geolocalizando %s', address);
  try {
      location = googleGeocoder.geocode(address);
      } 
  catch (e) {
    msg = e.message;
    Logger.log('Error Google %s', msg);
    location = {'status': 'SCRIPT ERROR'};
  }
  
  if (location.status == 'SCRIPT ERROR') {
    insertDataIntoSheet(cells, row, [
      [foundAddressColumn, ''], [latColumn, ''], [lngColumn, ''], [qualityColumn, 'FAILED SCRIPT: ' + msg], [sourceColumn, 'Google']
    ]);

    return 1;
  }
  
  if (location.status !== 'OK') {
    insertDataIntoSheet(cells, row, [
      [foundAddressColumn, ''], [latColumn, ''], [lngColumn, ''], [qualityColumn, 'No Match'], [sourceColumn, 'Google']
    ]);

    return 1;
  }

  lat = location['results'][0]['geometry']['location']['lat'];
  lng = location['results'][0]['geometry']['location']['lng'];
  foundAddress = location['results'][0]['formatted_address'];

  var quality;
  if (location['results'][0]['partial_match']) {
    quality = 'Partial Match';
  } else {
    quality = 'Match';
  }

  insertDataIntoSheet(cells, row, [
    [foundAddressColumn, foundAddress],
    [latColumn, lat],
    [lngColumn, lng],
    [qualityColumn, quality],
    [sourceColumn, 'Google']
  ]);

  return 0;
}


/**
 * Sets cells from a 'row' to values in data
 */
function insertDataIntoSheet(cells, row, data) {
  for (d in data) {
    cells.getCell(row, data[d][0]).setValue(data[d][1]);
  }
}

function censusAddressToPosition() {
  geocode('US Census');
}

function googleAddressToPosition() {
  geocode('Google');
}

function onOpen() {
  ui.createMenu('Geocodificar')
   .addItem('con Google (limite de 1000 por día)', 'googleAddressToPosition')
   // TODO agregar uno para OSM
   .addToUi();
}