const { stringify } = require('wkt');
const Excel = require('exceljs');
const jsonfile = require('./temp.json');

const run = async () => {

  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  worksheet.columns = [
    { header: 'name', key: 'name'},
    { header: 'description', key: 'description'},
    { header: 'coordinates', key: 'coordinates'},
    { header: 'location', key: 'location'},
    { header: 'type', key: 'type' }
  ];

  const features = jsonfile['features'];

  for (let k = 0; k < features.length; k++) {
    const name = features[k]['properties']['name'] || '';
    const description = features[k]['properties']['description'] || '';
    const type = features[k]['geometry']['type'] || '';

    const wkt = transformToWkt(features[k]['geometry']);

    const data = {
      name,
      description,
      type
    };

    if (type == 'Point' || type == 'LineString') {
      data['coordinates'] = wkt;
    } else if (type == 'Polygon') {
      data['location'] = wkt;
    }

    worksheet.addRow(data);
  }
  
  await workbook.xlsx.writeFile('Master Site Layout.xlsx');

  console.log("File is written");
};

/**
 * @param {object} objLocation 
 */
const transformToWkt = (objLocation) => {
  const data = {};
  data['type'] = objLocation['type'];
  
  if (objLocation['type'] === 'Polygon') {
    data['coordinates'] = cleanCoordinatePoligonArray(objLocation['coordinates']);
  } else if (objLocation['type'] === 'GeometryCollection') {
    data['geometries'] = cleanCoordinateGCArray(objLocation['geometries']);
  } else if (objLocation['type'] === 'Point') {
    data['coordinates'] = transformPointTo2D(objLocation['coordinates']);
  } else if (objLocation['type'] === 'LineString') {
    data['coordinates'] = transformLineStringTo2D(objLocation['coordinates']);
  } else {
    // todo: handle other type...
    console.log('error bang...');
    throw new Error('other type not handle yet');
  }

  const wktData = stringify(data);
  // RAW DATA //
  // NO NEED TO CLEANUP
  // 1. Z STRING
  // 2. 3D POINTS
  // --------- //
  // const data = {};
  // data['type'] = objLocation['type'];
  // data['coordinates'] = objLocation['coordinates'];
  // const wktData = stringify(data);

  return wktData;
};


/**
 * @param {Array<Integer>} array - [1, 2, 3]
 */
const transformPointTo2D = (array) => {
  if (array && array.length > 0) {
    array.pop()
    return array;
  } else {
    return [];
  }
}


/**
 * @param {Array<Array>} array - [ [1, 1, 1], [2, 2, 2], [3, 3, 3], ...]
 */
const transformLineStringTo2D = (array) => {
  if (array && array.length > 0) {
    array.forEach(arr => arr.pop());
    return array;
  } else {
    return array;
  }
};


/**
 * @param {Array} array 
 */
const cleanCoordinatePoligonArray = (array) => {
  if (array && array.length > 0) {
    array.forEach(arrPol => arrPol.forEach(arr => arr.pop()));
    return array;
  } else {
    return [];
  }
};


const cleanCoordinateGCArray = (array) => {
  if (array !== undefined) {
    array.forEach(arr => cleanCoordinatePoligonArray(arr['coordinates']));
  }

  return array;
};


run();
