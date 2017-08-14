# Docx-merger

Javascript Library for Merging Docx file in NodeJS and Browser Environment.

## Purpose

 To merge docx file using javascript which has rich contents.

 The Library Preserves the Styles, Tables, Images, Bullets and Numberings of input files.

## Table of Contents

  1. [Installation](#installation)
  1. [Usage NodeJS](#usage-nodeJS)
  1. [Usage Browser](#usage-browser)
  1. [TODO](#todo)
  1. [Known Issues](#known-issues)


## Installation


  ```bash
  npm install docx-merger
  ```

**[Back to top](#table-of-contents)**

### Usage NodeJS

Read input files as binary and pass it to the `DocxMerger` constructor fuction as a array of files.

Then call the save function with first argument as `nodebuffer`, check the example below.

  ```javascript
  var DocxMerger = require('docx-merger');

  var fs = require('fs');
  var path = require('path');

  var file1 = fs
      .readFileSync(path.resolve(__dirname, 'template.docx'), 'binary');

  var file2 = fs
      .readFileSync(path.resolve(__dirname, 'template1.docx'), 'binary');

  var docx = new DocxMerger({},[file1,file2]);


  //SAVING THE DOCX FILE

  docx.save('nodebuffer',function (data) {
      // fs.writeFile("output.zip", data, function(err){/*...*/});
      fs.writeFile("output.docx", data, function(err){/*...*/});
  });
  ```

**[Back to top](#table-of-contents)**

### Usage Browser

  -
  ```javascript

  ```



  ```javascript


  ```

### TODO

  -

  **[Back to top](#table-of-contents)**

### Known Issues

  -

  **[Back to top](#table-of-contents)**