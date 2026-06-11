'use strict';
var fflate = require('fflate');
var assert = require('node:assert');

var DOMParser = require('@xmldom/xmldom').DOMParser;

function toU8(data) {
    if (data instanceof Uint8Array) return data; // includes Buffer
    if (typeof data === 'string') {
        var u8 = new Uint8Array(data.length);
        for (var i = 0; i < data.length; i++) u8[i] = data.charCodeAt(i) & 0xff;
        return u8;
    }
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    throw new TypeError('unsupported data type in test helper');
}

function unzip(data) {
    return fflate.unzipSync(toU8(data));
}

function partText(entries, name) {
    assert.ok(entries[name], 'missing zip entry: ' + name);
    return fflate.strFromU8(entries[name]);
}

// strip all tags -> visible text of a part (entities left as-is)
function extractText(entries, name) {
    return partText(entries, name).replace(/<[^>]+>/g, '');
}

var REQUIRED_PARTS = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml',
                      'word/styles.xml', 'word/_rels/document.xml.rels'];

function assertValidDocx(data) {
    var entries = unzip(data);
    REQUIRED_PARTS.forEach(function (name) {
        assert.ok(entries[name], 'output is missing required part: ' + name);
    });
    Object.keys(entries).forEach(function (name) {
        if (!/\.(xml|rels)$/.test(name)) return;
        var xml = partText(entries, name);
        var errors = [];
        var record = function (msg) { errors.push(name + ': ' + msg); };
        // object form works on both xmldom 0.1.x (Phase 0) and @xmldom/xmldom 0.8 (Phase 1+)
        new DOMParser({
            errorHandler: { warning: function () {}, error: record, fatalError: record }
        }).parseFromString(xml, 'text/xml');
        assert.deepStrictEqual(errors, [], 'output part is not well-formed XML: ' + errors.join('; '));
    });
    return entries;
}

module.exports = { unzip: unzip, partText: partText, extractText: extractText, assertValidDocx: assertValidDocx };
