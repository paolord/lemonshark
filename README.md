lemonshark
==========
[![Build Status](https://semaphoreci.com/api/v1/paolord/lemonshark/branches/master/badge.svg)](https://semaphoreci.com/paolord/lemonshark) [![npm](https://img.shields.io/npm/dt/lemonshark.svg)](https://www.npmjs.com/package/lemonshark) [![Codecov](https://img.shields.io/codecov/c/github/paolord/lemonshark.svg?maxAge=2592000)](https://codecov.io/gh/paolord/lemonshark)

Simple spreadsheet reader module for node.js using iterator pattern.

IMPORTANT NOTE: not a production npm module.

## Installation

  npm install lemonshark --save

## Usage

```
  var reader = require('lemonshark').Reader;

  reader.open('./file.xlsx');

  // iterate over each sheet in workbook
  reader.sheetIterator().each(function(sheet) {
    var name = sheet.name; // sheet name
    var rows = sheet.rows; // sheet rows as array
  });

  reader.open('./another_file.xlsx');

  var iterator = reader.sheetIterator();

  var first_sheet = iterator.currentSheet();

  iterator.next(); // move to next sheet


  var second_sheet = iterator.currentSheet();

  iterator.prev(); // move to previous sheet

  iterator.rewind(); // return to initial sheet
```

## Tests

``` bash
  npm test
```
