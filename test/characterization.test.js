'use strict';
var test = require('node:test');
var assert = require('node:assert');
var fs = require('fs');
var path = require('path');

var DocxMerger = require('../src/index.js');
var DocxMergerV1 = require('docx-merger-v1');
var build = require('./helpers/build-docx.js');
var inspect = require('./helpers/inspect.js');

var tplA = fs.readFileSync(path.join(__dirname, '..', 'example', 'template.docx'), 'binary');
var tplB = fs.readFileSync(path.join(__dirname, '..', 'example', 'template1.docx'), 'binary');

function mergeWith(Ctor, files, options) {
    var out;
    new Ctor(options || {}, files).save('uint8array', function (d) { out = d; });
    return out;
}

// Tag-stripping concatenates adjacent runs and leaves inter-tag whitespace, so a
// standalone part and the same part inside a merge differ only in whitespace.
// Collapse it before comparing visible text.
function norm(s) {
    return s.replace(/\s+/g, '');
}

test('merging the example templates produces a structurally valid docx', function () {
    var entries = inspect.assertValidDocx(mergeWith(DocxMerger, [tplA, tplB]));
    var text = norm(inspect.extractText(entries, 'word/document.xml'));
    var a = norm(inspect.extractText(inspect.unzip(tplA), 'word/document.xml'));
    var b = norm(inspect.extractText(inspect.unzip(tplB), 'word/document.xml'));
    assert.ok(text.indexOf(a.slice(0, 40)) !== -1, 'text of first input missing from output');
    assert.ok(text.indexOf(b.slice(0, 40)) !== -1, 'text of second input missing from output');
    assert.ok(text.indexOf(a.slice(0, 40)) < text.indexOf(b.slice(0, 40)), 'input order not preserved');
});

test('page break is inserted between files by default and omitted with pageBreak:false', function () {
    var withBreak = inspect.partText(inspect.unzip(mergeWith(DocxMerger, [tplA, tplB])), 'word/document.xml');
    var without = inspect.partText(inspect.unzip(mergeWith(DocxMerger, [tplA, tplB], { pageBreak: false })), 'word/document.xml');
    assert.ok(/<w:br w:type="page"\/>/.test(withBreak));
    assert.ok(!/<w:br w:type="page"\/>/.test(without));
});

test('synthetic two-file merge keeps both texts', function () {
    var f1 = build.buildDocx({ body: build.para('ALPHA-CONTENT-ONE') });
    var f2 = build.buildDocx({ body: build.para('BETA-CONTENT-TWO') });
    var entries = inspect.assertValidDocx(mergeWith(DocxMerger, [f1, f2]));
    var text = inspect.extractText(entries, 'word/document.xml');
    assert.ok(text.indexOf('ALPHA-CONTENT-ONE') !== -1);
    assert.ok(text.indexOf('BETA-CONTENT-TWO') !== -1);
});

test('save callback is invoked synchronously (compat contract)', function () {
    var called = false;
    new DocxMerger({}, [tplA, tplB]).save('uint8array', function () { called = true; });
    assert.strictEqual(called, true, 'save() callback must fire before save() returns');
});

// ---- differential vs published 1.2.2 -------------------------------------

test('DIFFERENTIAL: body text identical to docx-merger@1.2.2 on example templates', function () {
    var current = inspect.extractText(inspect.unzip(mergeWith(DocxMerger, [tplA, tplB])), 'word/document.xml');
    var v1 = inspect.extractText(inspect.unzip(mergeWith(DocxMergerV1, [tplA, tplB])), 'word/document.xml');
    assert.strictEqual(current, v1);
});

test('DIFFERENTIAL: output part list is a superset of v1.2.2 part list', function () {
    var current = Object.keys(inspect.unzip(mergeWith(DocxMerger, [tplA, tplB]))).filter(function (n) { return !/\/$/.test(n); }).sort();
    var v1 = Object.keys(inspect.unzip(mergeWith(DocxMergerV1, [tplA, tplB]))).filter(function (n) { return !/\/$/.test(n); }).sort();
    v1.forEach(function (name) {
        assert.ok(current.indexOf(name) !== -1, 'part present in v1 output but missing now: ' + name);
    });
});
