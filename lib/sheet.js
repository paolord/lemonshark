var xmlreader = require('xmlreader');
var _ = require('underscore');

var sharedStrings = [];

function initializeSharedStrings(zip) {
  sharedStrings = [];
  xmlreader.read(zip.readAsText('xl/sharedStrings.xml'), function(err, res){
    if(err) return console.log(err);

    res.sst.si.each(function( i, si){
      sharedStrings.push(si.t.text());
    });
  });
}

function sharedStringLookup(id) {
  return sharedStrings[id];
}


exports.load = function(zip, sheetMeta, sheetData) {
  var _rows = [];

  if(sheetData.row === undefined){
    return {
      name: sheetMeta.name,
      id: sheetMeta.id,
      rows: []
    };
  }

  initializeSharedStrings(zip);

  sheetData.row.each(function(i, row) {
    var cols = [];

    row.c.each(function(i, col){
      if(col.attributes().t !== undefined){
        cols.push( sharedStringLookup(col.v.text()) );
      } else {
        cols.push( col.v.text() );
      }
    });
    //console.log(cols);
    _rows.push(cols);
  });

  return {
    name: sheetMeta.name,
    id: sheetMeta.id,
    rows: _rows
  };
};
