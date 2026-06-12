'use strict';
var DocxMerger = require('../src/index.js');
var build = require('../test/helpers/build-docx.js');

var styles = '';
var body = '';
for (var i = 0; i < 300; i++) {
    styles += '<w:style w:type="paragraph" w:styleId="S' + i + '"><w:name w:val="S' + i + '"/></w:style>';
}
for (var j = 0; j < 2000; j++) {
    body += '<w:p><w:pPr><w:pStyle w:val="S' + (j % 300) + '"/></w:pPr><w:r><w:t>para ' + j + '</w:t></w:r></w:p>';
}
var f1 = build.buildDocx({ styles: styles, body: body });
var f2 = build.buildDocx({ styles: styles, body: body });

var t0 = Date.now();
new DocxMerger({}, [f1, f2]).save('uint8array', function () {});
console.log('merge of 2x(300 styles, 2000 paragraphs): ' + (Date.now() - t0) + ' ms');
