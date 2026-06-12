'use strict';
var DOMParser = require('@xmldom/xmldom').DOMParser;
var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var xmlUtils = require('./xml-utils');

// Hyperlink relationship IDs are only referenced from each file's own
// document.xml, so we can safely rename them per file to avoid the Id collision
// that mergeRelations would otherwise resolve by silently dropping the duplicate.
var prepareHyperlinks = function (files) {
    var serializer = new XMLSerializer();

    files.forEach(function (zip, index) {
        if (index === 0) return; // base file keeps its IDs
        var relString = zip.getText('word/_rels/document.xml.rels');
        if (relString === null) return;
        var xml = new DOMParser().parseFromString(relString, 'text/xml');
        var rels = xml.getElementsByTagName('Relationship');
        var renames = [];

        for (var i = 0; i < rels.length; i++) {
            var rel = rels.item(i);
            if (/\/hyperlink$/.test(rel.getAttribute('Type'))) {
                var oldId = rel.getAttribute('Id');
                var newId = oldId + '_hl' + index;
                rel.setAttribute('Id', newId);
                renames.push([oldId, newId]);
            }
        }
        if (!renames.length) return;

        zip.setText('word/_rels/document.xml.rels',
            xmlUtils.replaceFrom(relString, '<Relationships', serializer.serializeToString(xml.documentElement)));

        var docString = zip.getText('word/document.xml');
        renames.forEach(function (pair) {
            docString = docString.replace(
                new RegExp('r:id="' + xmlUtils.escapeRegExp(pair[0]) + '"', 'g'),
                'r:id="' + pair[1] + '"');
        });
        zip.setText('word/document.xml', docString);
    });
};

module.exports = { prepareHyperlinks: prepareHyperlinks };
