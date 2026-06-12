'use strict';
var test = require('node:test');
var assert = require('node:assert');
var DocxMerger = require('../src/index.js');
var build = require('./helpers/build-docx.js');
var inspect = require('./helpers/inspect.js');

function numberedDoc(marker) {
    return build.buildDocx({
        numbering:
            '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
            '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>',
        body: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>' + marker + '</w:t></w:r></w:p>'
    });
}

function mergeToU8(files) {
    var out;
    new DocxMerger({}, files).save('uint8array', function (d) { out = d; });
    return out;
}

test('A8: every numId referenced in the merged body exists in merged numbering.xml (#56, #43)', function () {
    var entries = inspect.assertValidDocx(mergeToU8([numberedDoc('first-list'), numberedDoc('second-list')]));
    var numbering = inspect.partText(entries, 'word/numbering.xml');
    var doc = inspect.partText(entries, 'word/document.xml');

    var refs = (doc.match(/<w:numId w:val="(\d+)"/g) || []).map(function (m) { return m.match(/\d+/)[0]; });
    assert.ok(refs.length >= 2, 'expected two numbered paragraphs in body, got ' + refs.length);
    refs.forEach(function (id) {
        assert.ok(numbering.indexOf('<w:num w:numId="' + id + '"') !== -1,
            'body references numId ' + id + ' which is not defined in numbering.xml');
    });
    // the two files' lists must stay distinct
    assert.notStrictEqual(refs[0], refs[refs.length - 1], 'both files were collapsed onto one numId');
    // every w:num must point at an existing w:abstractNum
    var absDefs = (numbering.match(/w:abstractNumId="(\d+)"/g) || []);
    (numbering.match(/<w:abstractNumId w:val="(\d+)"/g) || []).forEach(function (m) {
        var id = m.match(/\d+/)[0];
        assert.ok(absDefs.indexOf('w:abstractNumId="' + id + '"') !== -1, 'dangling abstractNumId ' + id);
    });
});
