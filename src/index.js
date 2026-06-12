'use strict';
var ZipArchive = require('./zip-archive');
var xmlUtils = require('./xml-utils');

var Style = require('./merge-styles');
var Media = require('./merge-media');
var RelContentType = require('./merge-relations-and-content-type');
var bulletsNumbering = require('./merge-bullets-numberings');
var Hyperlinks = require('./merge-hyperlinks');

var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function u8ToBinaryString(u8) {
    var s = '';
    for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
    return s;
}

function u8ToBase64(u8) {
    if (typeof Buffer !== 'undefined') return Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength).toString('base64');
    return btoa(u8ToBinaryString(u8));
}

function convertOutput(u8, type) {
    switch (type) {
        case 'uint8array': return u8;
        case 'arraybuffer': return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
        case 'nodebuffer': return Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength);
        case 'blob': return new Blob([u8], { type: DOCX_MIME });
        case 'string':
        case 'binarystring': return u8ToBinaryString(u8);
        case 'base64': return u8ToBase64(u8);
        default: throw new Error('docx-merger: unknown save type "' + type + '"');
    }
}

function DocxMerger(options, files) {

    this._body = [];
    this._header = [];
    this._footer = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    this._files = [];
    var self = this;
    (files || []).forEach(function(file) {
        self._files.push(new ZipArchive(file));
    });
    this._contentTypes = {};

    this._media = {};
    this._rel = {};

    this._builder = this._body;

    this.insertPageBreak = function() {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        this._builder.push(pb);
    };

    this.insertRaw = function(xml) {

        this._builder.push(xml);
    };

    this.mergeBody = function(files) {

        var self = this;
        this._builder = this._body;

        RelContentType.mergeContentTypes(files, this._contentTypes);
        Media.prepareMediaFiles(files, this._media);
        Hyperlinks.prepareHyperlinks(files);
        RelContentType.mergeRelations(files, this._rel);

        var numberingMaps = bulletsNumbering.prepareNumbering(files);
        bulletsNumbering.mergeNumbering(files, this._numbering);

        Style.prepareStyles(files, this._style, numberingMaps);
        Style.mergeStyles(files, this._style);

        files.forEach(function(zip, index) {
            var xml = zip.getText("word/document.xml");
            xml = xml.substring(xml.indexOf("<w:body>") + 8);
            xml = xml.substring(0, xml.indexOf("</w:body>"));
            xml = xml.substring(0, xml.lastIndexOf("<w:sectPr"));

            self.insertRaw(xml);
            if (self._pageBreak && index < files.length-1)
                self.insertPageBreak();
        });
    };

    this.save = function(type, callback) {

        var zip = this._files[0];

        var xml = zip.getText("word/document.xml");
        var startIndex = xml.indexOf("<w:body>") + 8;
        var endIndex = xml.lastIndexOf("<w:sectPr");
        if (endIndex === -1) endIndex = xml.lastIndexOf("</w:body>");

        xml = xmlUtils.replaceBetween(xml, startIndex, endIndex, this._body.join(''));

        RelContentType.generateContentTypes(zip, this._contentTypes);
        Media.copyMediaFiles(zip, this._media, this._files);
        RelContentType.generateRelations(zip, this._rel);
        bulletsNumbering.generateNumbering(zip, this._numbering);
        Style.generateStyles(zip, this._style);

        zip.setText("word/document.xml", xml);

        var data = convertOutput(zip.generate(), type);
        if (typeof callback === 'function') {
            callback(data);            // synchronous, exactly like jszip 2 — compat contract
            return;
        }
        return Promise.resolve(data);  // additive: promise API when no callback given
    };


    if (this._files.length > 0) {

        this.mergeBody(this._files);
    }
}


module.exports = DocxMerger;
