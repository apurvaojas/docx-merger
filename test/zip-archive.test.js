'use strict';
var test = require('node:test');
var assert = require('node:assert');
var ZipArchive = require('../src/zip-archive.js');
var build = require('./helpers/build-docx.js');

var binstr = build.buildDocx({ body: build.para('zip-wrapper-test') });

test('accepts binary string, Buffer, Uint8Array, ArrayBuffer', function () {
    var u8 = new Uint8Array(binstr.length);
    for (var i = 0; i < binstr.length; i++) u8[i] = binstr.charCodeAt(i) & 0xff;
    [binstr, Buffer.from(u8), u8, u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength)].forEach(function (input) {
        var zip = new ZipArchive(input);
        assert.ok(zip.getText('word/document.xml').indexOf('zip-wrapper-test') !== -1);
    });
});

test('getText returns null for missing entries', function () {
    assert.strictEqual(new ZipArchive(binstr).getText('word/numbering.xml'), null);
});

test('setText/getBytes/setBytes/names round-trip', function () {
    var zip = new ZipArchive(binstr);
    zip.setText('word/custom.xml', '<a/>');
    zip.setBytes('word/media/image1.png', build.TINY_PNG);
    assert.strictEqual(zip.getText('word/custom.xml'), '<a/>');
    assert.deepStrictEqual(zip.getBytes('word/media/image1.png'), build.TINY_PNG);
    assert.ok(zip.names().indexOf('word/media/image1.png') !== -1);
});

test('generate produces a zip that re-opens with the same content', function () {
    var zip = new ZipArchive(binstr);
    zip.setText('word/document.xml', zip.getText('word/document.xml').replace('zip-wrapper-test', 'MUTATED'));
    var out = new ZipArchive(zip.generate());
    assert.ok(out.getText('word/document.xml').indexOf('MUTATED') !== -1);
});

test('throws a descriptive error on unsupported input', function () {
    assert.throws(function () { new ZipArchive(42); }, /docx-merger/);
});
