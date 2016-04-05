var AdmZip = require('adm-zip');
var xmlreader = require('xmlreader');
var _ = require('underscore');

var sheetClass = require('./sheet.js');

var zipfile = "";
var index = 0;
var sheets = [];

module.exports = function(filename, sheetnames) {
  sheets = sheetnames;
  zipfile = filename;

  return sheetIterator;
};

var sheetIterator = {};

sheetIterator.each = function(callback) {

  var zip = new AdmZip(zipfile);
  var sheet = {};
//  var currentSheet = sheets[index];
  //var _sheetIterator = sheetIterator;
  _.forEach(sheets , function(currentSheet) {
    xmlreader.read(zip.readAsText('xl/worksheets/sheet'+ currentSheet.id +'.xml'), function(err, res){
      if(err) return console.log(err);
      sheet = sheetClass.load(zip, currentSheet, res.worksheet.sheetData);
    });

    //_sheetIterator.next();
    callback(sheet);
  });
  //sheetIterator = _sheetIterator;
  //sheetIterator.rewind();
};

sheetIterator.currentSheet = function() {
  return sheets[index];
};

sheetIterator.next = function() {
  if(index < sheets.length){
    index = index + 1;
  }
};

sheetIterator.prev = function() {
  if(index > 0){
    index = index - 1;
  }
};

sheetIterator.rewind = function(){
  index = 0;
};
