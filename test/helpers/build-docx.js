'use strict';
// Builds minimal, Word-valid .docx files in memory and returns them as a binary
// string (the documented `fs.readFileSync(path, 'binary')` input format). All parts
// the merger requires are always present; tests for missing-part handling supply
// their own broken inputs.
var fflate = require('fflate');

var W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
           'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';

function buildDocx(opts) {
    opts = opts || {};
    var entries = {};

    entries['[Content_Types].xml'] = fflate.strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        (opts.contentTypeDefaults || '') +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        (opts.numbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '') +
        '</Types>');

    entries['_rels/.rels'] = fflate.strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>');

    entries['word/document.xml'] = fflate.strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document ' + W_NS + '><w:body>' +
        (opts.body || '<w:p><w:r><w:t>placeholder</w:t></w:r></w:p>') +
        '<w:sectPr/></w:body></w:document>');

    entries['word/styles.xml'] = fflate.strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:styles ' + W_NS + '><w:docDefaults/>' +
        (opts.styles || '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>') +
        '</w:styles>');

    entries['word/_rels/document.xml.rels'] = fflate.strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        (opts.numbering ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : '') +
        (opts.rels || '') +
        '</Relationships>');

    if (opts.numbering) {
        entries['word/numbering.xml'] = fflate.strToU8(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:numbering ' + W_NS + '>' + opts.numbering + '</w:numbering>');
    }

    if (opts.media) {
        Object.keys(opts.media).forEach(function (name) {
            entries['word/media/' + name] = opts.media[name];
        });
    }
    if (opts.extraEntries) {
        Object.keys(opts.extraEntries).forEach(function (name) {
            entries[name] = fflate.strToU8(opts.extraEntries[name]);
        });
    }

    return u8ToBinaryString(fflate.zipSync(entries));
}

function u8ToBinaryString(u8) {
    var s = '';
    for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
    return s;
}

function para(text) {
    return '<w:p><w:r><w:t xml:space="preserve">' + text + '</w:t></w:r></w:p>';
}

// 1x1 transparent PNG, for media fixtures
var TINY_PNG = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82
]);

module.exports = { buildDocx: buildDocx, para: para, TINY_PNG: TINY_PNG, u8ToBinaryString: u8ToBinaryString };
