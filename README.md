# Docx-merger

Javascript Library for Merging Docx file in NodeJS and Browser Environment.

## Purpose

 To merge docx file using javascript which has rich contents.

 The Library Preserves the Styles, Tables, Images, Bullets and Numberings of input files.

## Table of Contents

  1. [Installation](#installation)
  1. [Usage Nodejs](#usage-nodejs)
  1. [Usage Browser](#usage-browser)
  1. [TODO](#todo)
  1. [Known Issues](#known-issues)


## Installation


  ```bash
  npm install docx-merger
  ```

**[Back to top](#table-of-contents)**

### Usage Nodejs

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

  - Async Load files using `jszip-utils` and then call the callback in the innermost callback.
  - Call the save function wit first argument as `blob`.
  - Better use Promises instead of callbacks.
  - Callback causes callback hell issue.

###### Using Callback
  ```html
<html>
<script src="../dist/docx-merger.min.js"></script>
<script src="https://fastcdn.org/FileSaver.js/1.1.20151003/FileSaver.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip-utils/0.0.2/jszip-utils.min.js"></script>
<!--
Mandatory in IE 6, 7, 8 and 9.
-->
<!--[if IE]>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jszip-utils/0.0.2/jszip-utils-ie.min.js"></script>
<![endif]-->
<script>
    function loadFile(url,callback){
        JSZipUtils.getBinaryContent(url,callback);
    }
    loadFile("template.docx",function(error,file1){
        loadFile("template1.docx",function(error,file2){

            var docx = new DocxMerger({},[file1,file2]);

            docx.save('blob',function (data) {
                saveAs(data,"output.docx");
            });
        })
    })
</script>
</html>
  ```


###### Using Promise
  ```html
<html>
<script src="../dist/docx-merger.min.js"></script>
<script src="https://fastcdn.org/FileSaver.js/1.1.20151003/FileSaver.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip-utils/0.0.2/jszip-utils.min.js"></script>
<!--
Mandatory in IE 6, 7, 8 and 9.
-->
<!--[if IE]>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jszip-utils/0.0.2/jszip-utils-ie.min.js"></script>
<![endif]-->
<script>
    function loadFile(url,callback) {

        return new Promise(function (resolve, reject) {

            JSZipUtils.getBinaryContent(url, function (err, data) {
                if (err) reject(err);
                resolve(data);
            });
        });
    }
    Promise.all([loadFile("template.docx"), loadFile("template1.docx")]).then(function(files){

        var docx = new DocxMerger({},files);

        docx.save('blob',function (data) {
            saveAs(data,"output.docx");
        });

    },function (err) {
        alert(err);
    })
</script>
</html>

  ```

### TODO

  - CLI Support
  - Unit Tests
  - ES6 Convertions

  **[Back to top](#table-of-contents)**

### Known Issues

  - Microsoft Word in windows Shows some error due to numbering.

  **[Back to top](#table-of-contents)**