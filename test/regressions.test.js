'use strict';
var test = require('node:test');
var assert = require('node:assert');
var DocxMerger = require('../src/index.js');
var build = require('./helpers/build-docx.js');
var inspect = require('./helpers/inspect.js');

function mergeToU8(files, options) {
    var out;
    new DocxMerger(options || {}, files).save('uint8array', function (d) { out = d; });
    return out;
}

function count(haystack, needle) {
    return haystack.split(needle).length - 1;
}

test('A3: document text containing $&, $\', $` survives merging intact (#48, #18)', function () {
    // The base file carries a unique marker. The second file carries the $-tokens.
    // In the broken save(), the "$&" in file 2's text expands to the *matched base
    // body* (which contains the marker), so the marker is duplicated. A correct
    // splice inserts the tokens literally: the marker stays unique and the tokens
    // survive verbatim.
    var f1 = build.buildDocx({ body: build.para('FIRST_FILE_UNIQUE_MARKER') });
    var f2 = build.buildDocx({ body: build.para('Profit $&amp; Loss and $` and $\' tokens') });
    var entries = inspect.assertValidDocx(mergeToU8([f1, f2]));
    var text = inspect.extractText(entries, 'word/document.xml');

    assert.strictEqual(count(text, 'FIRST_FILE_UNIQUE_MARKER'), 1,
        'base body was duplicated by $-pattern expansion');
    // extractText strips tags but leaves entities, so "&" stays "&amp;"
    assert.ok(text.indexOf("Profit $&amp; Loss and $` and $' tokens") !== -1,
        'second file $-tokens were mangled: ' + text);
});
