
var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var DOMParser = require('@xmldom/xmldom').DOMParser;
var xmlUtils = require('./xml-utils');


var mergeContentTypes = function(files, _contentTypes) {


    files.forEach(function(zip) {
        var xmlString = zip.getText("[Content_Types].xml");
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Types')[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var contentType = childNodes[node].getAttribute('ContentType');
                if (!_contentTypes[contentType])
                    _contentTypes[contentType] = childNodes[node].cloneNode();
            }
        }

    });
};

var mergeRelations = function(files, _rel) {

    files.forEach(function(zip) {
        var xmlString = zip.getText("word/_rels/document.xml.rels");
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var Id = childNodes[node].getAttribute('Id');
                if (!_rel[Id])
                    _rel[Id] = childNodes[node].cloneNode();
            }
        }

    });
};

var generateContentTypes = function(zip, _contentTypes) {
    // body...
    var xmlString = zip.getText("[Content_Types].xml");
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    for (var node in _contentTypes) {
        types.appendChild(_contentTypes[node]);
    }

    xmlString = xmlUtils.replaceFrom(xmlString, "<Types", serializer.serializeToString(types));

    zip.setText("[Content_Types].xml", xmlString);
};

var generateRelations = function(zip, _rel) {
    // body...
    var xmlString = zip.getText("word/_rels/document.xml.rels");
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    for (var node in _rel) {
        types.appendChild(_rel[node]);
    }

    xmlString = xmlUtils.replaceFrom(xmlString, "<Relationships", serializer.serializeToString(types));

    zip.setText("word/_rels/document.xml.rels", xmlString);
};


module.exports = {
    mergeContentTypes: mergeContentTypes,
    mergeRelations: mergeRelations,
    generateContentTypes: generateContentTypes,
    generateRelations: generateRelations
};
