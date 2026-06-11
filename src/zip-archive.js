'use strict';
var fflate = require('fflate');

function toUint8Array(input) {
    if (input instanceof Uint8Array) return input; // includes Buffer
    if (typeof input === 'string') {
        var u8 = new Uint8Array(input.length);
        for (var i = 0; i < input.length; i++) u8[i] = input.charCodeAt(i) & 0xff;
        return u8;
    }
    if (input instanceof ArrayBuffer) return new Uint8Array(input);
    throw new TypeError('docx-merger: unsupported file input; expected binary string, Buffer, Uint8Array or ArrayBuffer');
}

function ZipArchive(input) {
    try {
        this._entries = fflate.unzipSync(toUint8Array(input));
    } catch (e) {
        if (e instanceof TypeError) throw e;
        throw new Error('docx-merger: input is not a valid docx/zip file (' + e.message + ')');
    }
}

ZipArchive.prototype.getText = function (name) {
    var data = this._entries[name];
    return data ? fflate.strFromU8(data) : null;
};

ZipArchive.prototype.setText = function (name, text) {
    this._entries[name] = fflate.strToU8(text);
};

ZipArchive.prototype.getBytes = function (name) {
    return this._entries[name] || null;
};

ZipArchive.prototype.setBytes = function (name, bytes) {
    this._entries[name] = bytes;
};

ZipArchive.prototype.remove = function (name) {
    delete this._entries[name];
};

ZipArchive.prototype.names = function () {
    return Object.keys(this._entries).filter(function (n) { return n.slice(-1) !== '/'; });
};

// returns Uint8Array; level 4 matches the DEFLATE level on master (commit af8f4b6)
ZipArchive.prototype.generate = function () {
    return fflate.zipSync(this._entries, { level: 4 });
};

module.exports = ZipArchive;
