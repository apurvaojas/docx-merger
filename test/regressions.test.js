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

test('A4: style IDs with regex metacharacters merge correctly (no crash, ref renamed)', function () {
    var styleId = 'Title(Main'; // unbalanced paren -> an unescaped RegExp throws SyntaxError
    var styles = '<w:style w:type="paragraph" w:styleId="' + styleId + '"><w:name w:val="x"/></w:style>';
    var f1 = build.buildDocx({
        styles: styles,
        body: '<w:p><w:pPr><w:pStyle w:val="' + styleId + '"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>'
    });
    var f2 = build.buildDocx({ body: build.para('plain') });
    var doc = inspect.partText(inspect.assertValidDocx(mergeToU8([f1, f2])), 'word/document.xml');
    assert.ok(doc.indexOf('w:val="' + styleId + '_0"') !== -1, 'style reference in body was not renamed: ' + doc);
});

function withImages(imgNames) {
    var media = {}, rels = '', body = '';
    imgNames.forEach(function (n, i) {
        media[n] = build.TINY_PNG;
        rels += '<Relationship Id="rImg' + i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + n + '"/>';
        body += '<w:p><w:r><w:t>img ' + n + '</w:t></w:r></w:p>';
    });
    return build.buildDocx({ media: media, rels: rels, body: body, contentTypeDefaults: '<Default Extension="png" ContentType="image/png"/>' });
}

test('A6: multi-digit media names do not collide or duplicate (#55)', function () {
    var f1 = withImages(['image1.png', 'image10.png']);
    var f2 = withImages(['image1.png']);
    var entries = inspect.assertValidDocx(mergeToU8([f1, f2]));
    var mediaNames = Object.keys(entries).filter(function (n) { return /^word\/media\//.test(n); });
    assert.strictEqual(mediaNames.length, 3,
        'expected exactly 3 media files, got: ' + mediaNames.join(', '));
});
