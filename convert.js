const { stringify } = require('wkt');
const Excel = require('exceljs');
const jsonfile = require('./temp.json');

const run = async () => {

  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  worksheet.columns = [
    { header: 'No_Bid', key: 'noBid'},
    { header: 'location', key: 'location'}
  ];

  const features = jsonfile['features'];
  for (let k=0; k<features.length; k++) {
    const noBid = features[k]['properties']['No_Bid'];
    const location = transformLocation(features[k]['geometry']);
    const data = {
      noBid,
      location
    };
    worksheet.addRow(data);
  }
  
  await workbook.xlsx.writeFile('MDA_bidang2020_Full.xlsx');

  console.log("File is written");
};

/**
 * @param {object} objLocation 
 */
const transformLocation = (objLocation) => {
  if (objLocation['type'] === 'Polygon') {
    const data = {};
    data['type'] = objLocation['type'];
    data['coordinates'] = cleanCoordinatePoligonArray(objLocation['coordinates']);

    const wktData = stringify(data);
    return wktData;
  } else if (objLocation['type'] === 'GeometryCollection') {
    const data = {};
    data['type'] = objLocation['type'];
    data['geometries'] = cleanCoordinateGCArray(objLocation['geometries']);
    
    const wktData = stringify(data);
    return wktData
  } else {
    // todo: handle other type...
    throw new Error('other type not handle yet');
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
