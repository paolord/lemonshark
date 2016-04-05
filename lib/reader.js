var AdmZip = require('adm-zip');
var xmlreader = require('xmlreader');

var filenameGlobal = "";

var sheetsMeta = [];
var sharedStrings = [];

exports.open = function(filename) {
  filenameGlobal = "";
  sheetsMeta = [];
  sharedStrings = [];

  filenameGlobal = filename;
  var zip = new AdmZip(filename);

  sheetHelper(zip);

  return sheetsMeta;
};

function sheetHelper(zip){
  _sheets = sheetsMeta;
  xmlreader.read(zip.readAsText('xl/workbook.xml'), function(err, res){
    if(err) return console.log(err);

    res.workbook.sheets.sheet.each(function( i, sheet){
      //_sheets.push( require('./Sheet.js')( zip, {name: sheet.attributes().name, id: sheet.attributes().sheetId}) );
      _sheets.push( {name: sheet.attributes().name, id: sheet.attributes().sheetId} );
    });
  });
  
  sheetsMeta = _sheets;
}

exports.sheetIterator = function() {
  return require('./sheetIterator.js')(filenameGlobal, sheetsMeta);
};
