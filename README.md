# Docx-merger

Javascript Library for Merging Docx file in NodeJS and Browser Environment.

## Purpose

 To merge docx file using javascript which has rich contents.

 The Library Preserves the Styles, Tables, Images, Bullets and Numberings of input files.

> **v2.0.0** is a security + correctness release with **no breaking changes to the
> documented API**. It removes the vulnerable `jszip@2` and abandoned `xmldom`
> dependencies and fixes long-standing output-corruption, media, hyperlink and
> numbering bugs. See [CHANGELOG.md](CHANGELOG.md) and
> [Migrating from 1.x](#migrating-from-1x) below.

## Table of Contents

  1. [Installation](#installation)
  1. [Usage Nodejs](#usage-nodejs)
  1. [Usage Browser](#usage-browser)
  1. [Input and output types](#input-and-output-types)
  1. [Migrating from 1.x](#migrating-from-1x)
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

#### Using a Promise (new in v2)

Call `save(type)` with no callback to get a Promise instead:

  ```javascript
  var data = await new DocxMerger({}, [file1, file2]).save('nodebuffer');
  fs.writeFileSync('output.docx', data);
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

### Input and output types

Each input file may be a binary string (`fs.readFileSync(path, 'binary')`), a
`Buffer`, a `Uint8Array`, or an `ArrayBuffer`.

The first argument to `save(type[, callback])` selects the output type:

| `type`                       | Output            | Typical use      |
| ---------------------------- | ----------------- | ---------------- |
| `nodebuffer`                 | `Buffer`          | Node, write file |
| `blob`                       | `Blob`            | Browser download |
| `uint8array`                 | `Uint8Array`      | generic binary   |
| `arraybuffer`                | `ArrayBuffer`     | generic binary   |
| `base64`                     | base64 `string`   | data URIs        |
| `string` / `binarystring`    | binary `string`   | legacy           |

With a callback, it is invoked synchronously. Without a callback, `save` returns
a `Promise` of the same value.

  **[Back to top](#table-of-contents)**

### Migrating from 1.x

For the documented usage above, **no code changes are required** — the
constructor and `save(type, callback)` behave the same. v2 additionally lets
`save(type)` return a Promise.

What changed under the hood (hence the major version bump):

  - `jszip@2` and `xmldom` were replaced internally with `fflate` and
    `@xmldom/xmldom`; output bytes differ (media parts are renamed, numbering is
    renumbered) but documents render identically or better.
  - Invalid input now throws a descriptive `Error` instead of an opaque
    `TypeError`.

  **[Back to top](#table-of-contents)**

### TODO

  - CLI Support
  - ES6 Conversion / TypeScript source
  - Header & footer merging

  **[Back to top](#table-of-contents)**

### Known Issues

  - List numbering across merged files (the old "Word found unreadable content"
    prompt) is **fixed in v2.0.0**.
  - Headers and footers are not yet merged — only the first document's are kept.

  **[Back to top](#table-of-contents)**