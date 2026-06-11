'use strict';
var test = require('node:test');
var assert = require('node:assert');
var DocxMerger = require('../src/index.js');
var build = require('./helpers/build-docx.js');
var inspect = require('./helpers/inspect.js');

function fixture() {
    return [build.buildDocx({ body: build.para('one') }), build.buildDocx({ body: build.para('two') })];
}

test('save("nodebuffer") yields a Buffer that is a valid docx', function () {
    new DocxMerger({}, fixture()).save('nodebuffer', function (data) {
        assert.ok(Buffer.isBuffer(data));
        inspect.assertValidDocx(data);
    });
});

test('save("uint8array") and save("arraybuffer")', function () {
    new DocxMerger({}, fixture()).save('uint8array', function (d) { assert.ok(d instanceof Uint8Array); });
    new DocxMerger({}, fixture()).save('arraybuffer', function (d) { assert.ok(d instanceof ArrayBuffer); });
});

test('save("base64") and save("binarystring") round-trip', function () {
    new DocxMerger({}, fixture()).save('base64', function (d) {
        inspect.assertValidDocx(new Uint8Array(Buffer.from(d, 'base64')));
    });
    new DocxMerger({}, fixture()).save('binarystring', function (d) {
        assert.strictEqual(typeof d, 'string');
        inspect.assertValidDocx(d);
    });
});

test('save without callback returns a Promise (new in v2)', async function () {
    var data = await new DocxMerger({}, fixture()).save('uint8array');
    inspect.assertValidDocx(data);
});

test('save with unknown type throws a descriptive error', function () {
    assert.throws(function () { new DocxMerger({}, fixture()).save('csv', function () {}); }, /unknown save type/);
});
