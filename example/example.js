const DocxMerger = require('../dist/index.cjs');
const fs = require('fs');
const path = require('path');

const file1 = fs
    .readFileSync(path.resolve(__dirname, 'template.docx'), 'binary');

const file2 = fs
    .readFileSync(path.resolve(__dirname, 'template1.docx'), 'binary');

const docx = new DocxMerger({},[file1,file2]);


//SAVING THE DOCX FILE

docx.save('nodebuffer',function (data) {
    fs.writeFile("output.zip", data, function(err){});
    fs.writeFile("output.docx", data, function(err){});
});
