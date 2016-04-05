"use strict";

var expect = require('chai').expect;

var reader = require('../index.js').Reader;

describe('lemonshark XLSX', function(){
  it('should open XLSX file', function(){
    /*
    Dataset acquired from
    http://data.gov.ph/catalogue/dataset/number-of-tourism-establishment-and-rooms
    */
    var sheets = reader.open('./test/dot_num_estab_and_rooms.xlsx');

    expect(sheets).to.deep.equal([ { name: 'dot_num_estab_and_rooms', id: '1' }, { name: 'Sheet1', id: '2' } ]);
  });
  it('should retrieve the expected amount of rows for sheet1', function() {

    reader.open('./test/dot_num_estab_and_rooms.xlsx');

    reader.sheetIterator().each(function(sheet) {
      if(sheet.name === 'dot_num_estab_and_rooms'){
        expect(sheet.rows.length).to.be.equal(637 - 1);
      }
    });
  });
  it('should return an empty array for empty sheets', function() {
    reader.open('./test/dot_num_estab_and_rooms.xlsx');

    reader.sheetIterator().each(function(sheet) {
      if(sheet.name === 'sheet1'){
        expect(sheet.rows.length).to.be.equal(0);
      }
    });
  });
  describe('lemonshark iterator', function() {
    before(function() {
      reader.open('./test/dot_num_estab_and_rooms.xlsx');
    });
    it('should be current sheet', function() {
      expect(reader.sheetIterator().currentSheet().name).to.be.equal('dot_num_estab_and_rooms');
    });
    it('should move to next sheet using next()', function() {
      reader.sheetIterator().next();
      expect(reader.sheetIterator().currentSheet().name).to.be.equal('Sheet1');
    });
    it('should move to next sheet using prev()', function() {
      reader.sheetIterator().prev();
      expect(reader.sheetIterator().currentSheet().name).to.be.equal('dot_num_estab_and_rooms');
    });
    it('should rewind to starting sheet using rewind()', function() {
      reader.sheetIterator().next();
      reader.sheetIterator().rewind();
      expect(reader.sheetIterator().currentSheet().name).to.be.equal('dot_num_estab_and_rooms');
    });
  });
});
